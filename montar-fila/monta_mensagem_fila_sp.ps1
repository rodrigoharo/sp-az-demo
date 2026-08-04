param(
    [Parameter(Mandatory = $true)]
    [string]$SharePointUrl,

    [string]$TargetBlobRoot = "sp",

    [switch]$NoRecursive,

    [string]$OutputPath
)

if ($PSVersionTable.PSEdition -ne 'Core') {
    $fwd = [System.Collections.Generic.List[string]]::new()
    $fwd.AddRange([string[]]@('-ExecutionPolicy', 'ByPass', '-File', $MyInvocation.MyCommand.Path))
    foreach ($key in $PSBoundParameters.Keys) {
        $fwd.Add("-$key")
        if ($PSBoundParameters[$key] -isnot [switch]) {
            $fwd.Add([string]$PSBoundParameters[$key])
        }
    }
    & pwsh $fwd
    exit $LASTEXITCODE
}

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
    throw "Cmdlet Invoke-MgGraphRequest nao encontrado. Instale/importe Microsoft.Graph.Authentication e conecte com Connect-MgGraph."
}

if (-not (Get-Command Get-MgContext -ErrorAction SilentlyContinue)) {
    throw "Cmdlet Get-MgContext nao encontrado. Instale/importe Microsoft.Graph.Authentication."
}

$mgContext = Get-MgContext
if ($null -eq $mgContext) {
    Write-Host "Conectando no Microsoft Graph..."
    Connect-MgGraph -NoWelcome
    $mgContext = Get-MgContext
}
Write-Host "Conectado como: $($mgContext.Account) (Tenant: $($mgContext.TenantId))"

function ConvertTo-SharingToken {
    param([string]$Url)

    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Url)
    $b64 = [System.Convert]::ToBase64String($bytes).TrimEnd('=').Replace('/', '_').Replace('+', '-')
    return "u!$b64"
}

function ConvertTo-QueueSafePath {
    param([string]$Path)

    return ($Path.Trim('/') -replace '/+', '/')
}

function Remove-TextAccent {
    param([string]$Value)

    $normalized = $Value.Normalize([System.Text.NormalizationForm]::FormD)
    $builder = [System.Text.StringBuilder]::new()
    foreach ($char in $normalized.ToCharArray()) {
        $category = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($char)
        if ($category -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$builder.Append($char)
        }
    }
    return $builder.ToString()
}

function ConvertTo-AemFolderSegment {
    param([string]$Value)

    $text = (Remove-TextAccent $Value).ToLowerInvariant()
    $text = $text -replace '\s*-\s*', '-'
    $text = $text -replace '[^a-z0-9._-]', '-'
    $text = $text -replace '-{2,}', '-'
    $text = $text.Trim('-.')
    if ([string]::IsNullOrWhiteSpace($text)) {
        return "pasta"
    }
    return $text
}

function ConvertTo-AemBlobPrefix {
    param([string]$Path)

    $parts = (ConvertTo-QueueSafePath $Path).Split('/', [System.StringSplitOptions]::RemoveEmptyEntries)
    if ($parts.Count -eq 0) {
        return ""
    }

    $normalized = [System.Collections.Generic.List[string]]::new()
    $normalized.Add($parts[0])
    for ($idx = 1; $idx -lt $parts.Count; $idx++) {
        $normalized.Add((ConvertTo-AemFolderSegment $parts[$idx]))
    }
    return ($normalized -join '/')
}

function Resolve-SharePointFolderUrl {
    param([string]$Url)

    $Url = $Url.Trim()
    $uri = [System.Uri]$Url
    $hostname = $uri.Host

    if ($Url -match '/:[fFbBuUpPwW]:/') {
        Write-Host "Resolvendo link de compartilhamento..."
        $sharingToken = ConvertTo-SharingToken -Url $Url
        $item = Invoke-MgGraphRequest -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/shares/$sharingToken/driveItem" `
            -ErrorAction Stop

        $parent = $item.parentReference
        $driveId = $parent.driveId
        $drivePath = ($parent.path -replace '^/drives/[^/]+/root:', '').Trim('/')
        if (-not [string]::IsNullOrWhiteSpace($item.name)) {
            $drivePath = ConvertTo-QueueSafePath "$drivePath/$($item.name)"
        }

        $siteIdParts = ($parent.siteId -split ',')
        $siteHostname = if ($siteIdParts.Count -gt 0) { $siteIdParts[0] } else { $hostname }
        $site = Invoke-MgGraphRequest -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/sites/$($parent.siteId)" `
            -ErrorAction Stop
        $sitePath = ([System.Uri]$site.webUrl).AbsolutePath.TrimEnd('/')

        return @{
            Hostname = $siteHostname
            SitePath = $sitePath
            DriveId = $driveId
            FolderPath = $drivePath
            SiteName = ($sitePath -split '/')[-1]
        }
    }

    $serverRelPath = $null
    if (-not [string]::IsNullOrWhiteSpace($uri.Query)) {
        foreach ($pair in ($uri.Query.TrimStart('?') -split '&')) {
            $kv = $pair -split '=', 2
            if ($kv.Count -eq 2 -and $kv[0] -eq 'id') {
                $serverRelPath = [System.Uri]::UnescapeDataString($kv[1].Replace('+', ' '))
                break
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($serverRelPath)) {
        throw "URL sem parametro id=. Abra a pasta no SharePoint e copie a URL completa do navegador."
    }

    $cleanPath = $serverRelPath.TrimStart('/')
    if ($cleanPath -notmatch '^(sites|teams|personal)/([^/]+)/([^/]+)(?:/(.+))?$') {
        throw "Nao foi possivel interpretar o caminho do parametro id: $serverRelPath"
    }

    $sitePath = "/$($Matches[1])/$($Matches[2])"
    $libraryName = $Matches[3]
    $folderPath = if ($Matches[4]) { ConvertTo-QueueSafePath $Matches[4] } else { "" }
    $siteName = $Matches[2]

    Write-Host "Buscando site: $hostname$sitePath"
    $site = Invoke-MgGraphRequest -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}" `
        -ErrorAction Stop

    Write-Host "Buscando biblioteca: $libraryName"
    $drivesResp = Invoke-MgGraphRequest -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives" `
        -ErrorAction Stop

    $drive = $drivesResp.value | Where-Object {
        $_.name -eq $libraryName -or
        ([System.Uri]::UnescapeDataString($_.webUrl) -like "*/$libraryName")
    } | Select-Object -First 1

    if (-not $drive) {
        $available = ($drivesResp.value | Select-Object -ExpandProperty name) -join ', '
        throw "Biblioteca '$libraryName' nao encontrada. Disponiveis: $available"
    }

    return @{
        Hostname = $hostname
        SitePath = $sitePath
        DriveId = $drive.id
        FolderPath = $folderPath
        SiteName = $siteName
    }
}

$resolved = Resolve-SharePointFolderUrl -Url $SharePointUrl
$targetBlobPrefix = ConvertTo-AemBlobPrefix "$TargetBlobRoot/$($resolved.SiteName)/$($resolved.FolderPath)"

$message = [ordered]@{
    sharePointHostname = $resolved.Hostname
    sharePointSitePath = $resolved.SitePath
    sharePointDriveId = $resolved.DriveId
    sharePointFolderPath = $resolved.FolderPath
    targetBlobPrefix = $targetBlobPrefix
    recursive = -not $NoRecursive.IsPresent
    metadata = [ordered]@{
        origem = "sharepoint"
        site = $resolved.SiteName
    }
}

$json = $message | ConvertTo-Json -Depth 6

Write-Host ""
Write-Host "Mensagem para a fila demo-dam-migration-queue-folders:"
Write-Host ""
Write-Output $json

if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($OutputPath, $json + [Environment]::NewLine, $utf8NoBom)
    Write-Host ""
    Write-Host "OK -> $OutputPath"
}

