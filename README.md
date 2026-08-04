# SharePoint to Azure Blob to AEM DAM Migration Pipeline

Queue-based migration workflow that moves digital assets from SharePoint to Azure Blob Storage and prepares metadata for Adobe Experience Manager Assets bulk import.

## What It Does

- Reads an asset control workbook and converts eligible rows into queue messages.
- Resolves SharePoint folder links with Microsoft Graph.
- Discovers files recursively from source folders.
- Migrates files into Azure Blob Storage with normalized DAM-safe paths.
- Stores migration status in Azure Table Storage.
- Skips temporary or unsupported files such as zero-byte files and transient desktop files.
- Generates AEM bulk metadata CSV files from the migration status table.

## Architecture

```text
Workbook
  -> queue message builder
  -> Azure Queue: folders
  -> Azure Function: discover_sharepoint_folder
  -> Azure Queue: files
  -> Azure Function: migrate_file_worker
  -> Azure Blob Storage + Azure Table Storage
  -> AEM metadata CSV generator
  -> AEM Assets bulk import
```

More detail is available in [docs/architecture.md](docs/architecture.md).

## Repository Layout

```text
aem-metadata/      Generates AEM bulk metadata CSV files from migration status.
montar-fila/       Builds folder queue messages from a workbook.
normalizacao/      Local checks for DAM path normalization rules.
orquestrador/      Coordinates workbook parsing, queue publishing, and queue drain wait.
shared_code/       Shared Graph, Storage, config, and path normalization helpers.
validacao/         Workbook validation helpers.
function_app.py    Azure Functions queue triggers.
```

## Configuration

Copy `local.settings.json.example` to `local.settings.json` for local development and fill in your own demo resources:

```json
{
  "QUEUE_STORAGE_CONNECTION_STRING": "<queue-storage-connection-string>",
  "DEST_STORAGE_ACCOUNT_URL": "https://demodamstorage.blob.core.windows.net/",
  "DEST_CONTAINER_NAME": "demo-dam-migration",
  "QUEUE_NAME": "demo-dam-migration-queue-files",
  "FOLDER_QUEUE_NAME": "demo-dam-migration-queue-folders",
  "STATUS_TABLE_NAME": "MigrationStatus",
  "SHAREPOINT_TENANT_ID": "<tenant-id>",
  "SHAREPOINT_CLIENT_ID": "<graph-app-client-id>",
  "SHAREPOINT_CLIENT_SECRET": "<graph-app-client-secret>"
}
```

## Example Commands

Generate queue messages from a workbook:

```powershell
python .\montar-fila\monta_lote_planilha_sp.py `
  --workbook "C:\demo\DAM-ASSETS-BATCH-01.xlsx" `
  --sheet "DAM-ASSETS" `
  --start-row 5 `
  --filter-tag "BATCH-01" `
  --target-blob-root "sp/ativos" `
  --output-dir ".\tmp\batch-01" `
  --graph-auth function-app `
  --resource-group "rg-demo-dam" `
  --function-app-name "sp-dam-migration-worker"
```

Run the orchestrator and wait for queues to drain:

```powershell
python .\orquestrador\orquestra_lote_planilha.py `
  --workbook "C:\demo\DAM-ASSETS-BATCH-01.xlsx" `
  --sheet "DAM-ASSETS" `
  --start-row 5 `
  --filter-tag "BATCH-01" `
  --send-to-queue `
  --wait `
  --output-dir ".\tmp\batch-01"
```

Generate an AEM metadata CSV from the migration status table:

```powershell
python .\aem-metadata\gera_csv_bulk_aem.py `
  --aem-root "/content/dam/demo" `
  --blob-prefix "sp/ativos/24-25/backup/" `
  --strip-blob-prefix "sp/ativos" `
  --output-dir ".\tmp\aem-metadata" `
  --resource-group "rg-demo-dam" `
  --function-app-name "sp-dam-migration-worker"
```

## Operational Notes

- Keep `local.settings.json` local only.
- Review blob counts before using the generated AEM metadata CSV.
- Treat poison queue messages as retry and triage inputs.

## Tech Stack

- Python
- Azure Functions
- Azure Queue Storage
- Azure Blob Storage
- Azure Table Storage
- Microsoft Graph API
- Adobe Experience Manager Assets bulk metadata import
