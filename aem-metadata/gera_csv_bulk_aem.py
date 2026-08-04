#!/usr/bin/env python
from __future__ import annotations

import argparse
import csv
import json
import os
import re
import shutil
import subprocess
import sys
import unicodedata
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path, PurePosixPath
from typing import Any


DEFAULT_METADATA_COLUMNS = [
    ("demo:anoUso", "String"),
    ("demo:cdISA", "String"),
    ("demo:cdRede", "String"),
    ("demo:ciclo", "String"),
    ("demo:colecao", "String"),
    ("demo:caminhoBlobNormalizado", "String"),
    ("demo:caminhoOriginalSharePoint", "String"),
    ("demo:damRoot", "String"),
    ("demo:damCiclo", "String"),
    ("demo:damCategoria", "String"),
    ("demo:damCodigoColecao", "String"),
    ("demo:damOrigemConteudo", "String"),
    ("demo:damRegraTaxonomia", "String"),
    ("demo:damRetrancaMaterial", "String"),
    ("demo:damTipo", "String"),
    ("demo:nome", "String"),
    ("demo:nomeArquivoOriginal", "String"),
    ("demo:pastaTecnica", "String"),
    ("demo:produto", "String"),
    ("demo:qtdPaginas", "Long"),
    ("demo:referenciaColecaoLote", "String"),
    ("demo:retrancaAnterior", "String"),
    ("demo:saida", "String"),
    ("demo:segmento", "String"),
    ("demo:selo", "String"),
    ("demo:sharePointSourceUrl", "String"),
    ("demo:sintaxeNova", "String"),
    ("demo:statusProduto", "String"),
    ("demo:tagsFase1", "String"),
]

IDENTIFIER_COLUMNS = {"demo:cdISA", "demo:cdRede"}
INTEGER_COLUMNS = {"demo:qtdPaginas"}
EXCLUDED_AEM_METADATA_EXTENSIONS = {
    ".bmap",
    ".idlk",
    ".lst 2",
    ".textclipping",
}
DAM_CATEGORY_BY_SEGMENT = {
    "documentacao-geral": "documentacao-geral",
    "digital": "digital",
    "grafica": "grafica",
    "pnld": "pnld",
    "acessibilidade": "acessibilidade",
    "ativos-digital": "digital",
    "ativos-grafica": "grafica",
    "ativos-pnld": "pnld",
    "ativos-accessibility": "acessibilidade",
    "politicas-e-procedimentos": "Politicas-e-Procedimentos",
    "manuais-e-guias": "Manuais-e-Guias",
    "contratos-e-garantias": "Contratos-e-Garantias",
}


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Gera CSV de bulk metadata do AEM a partir da tabela de status da migracao."
    )
    parser.add_argument("--output-dir", required=True, help="Pasta de saida.")
    parser.add_argument(
        "--aem-root",
        required=True,
        help="Raiz no AEM onde os blobs foram importados. Ex: /content/dam/demo",
    )
    parser.add_argument(
        "--blob-prefix",
        action="append",
        default=[],
        help="Filtra apenas TargetBlobName que comecam com este prefixo. Pode ser informado mais de uma vez.",
    )
    parser.add_argument(
        "--strip-blob-prefix",
        default="",
        help="Remove este prefixo do TargetBlobName antes de montar assetPath.",
    )
    parser.add_argument(
        "--status",
        default="Succeeded",
        help="Status da tabela que entra no CSV. Padrao: Succeeded.",
    )
    parser.add_argument(
        "--resource-group",
        default="rg-demo-dam",
        help="Resource group da Function usada para ler app settings.",
    )
    parser.add_argument(
        "--function-app-name",
        default="sp-dam-migration-worker",
        help="Function App usada para ler STATUS_TABLE_NAME e connection strings.",
    )
    parser.add_argument(
        "--table-name",
        help="Nome da tabela. Se omitido, usa STATUS_TABLE_NAME da Function App.",
    )
    parser.add_argument(
        "--connection-string",
        help="Connection string do Storage. Se omitido, tenta STATUS_STORAGE_CONNECTION_STRING, AzureWebJobsStorage e Managed Identity.",
    )
    parser.add_argument(
        "--storage-account-url",
        help="URL da storage account. Ex: https://demodamstorage.blob.core.windows.net/",
    )
    parser.add_argument(
        "--include-empty-metadata",
        action="store_true",
        help="Inclui arquivos migrados sem AemMetadataJson, deixando colunas de metadados vazias.",
    )
    parser.add_argument(
        "--max-rows",
        type=int,
        help="Limita a quantidade de linhas no CSV, util para teste.",
    )
    parser.add_argument(
        "--upload-to-blob",
        action="store_true",
        help="Tambem grava o CSV gerado no container de destino da migracao.",
    )
    parser.add_argument(
        "--output-blob-name",
        help="Nome do blob para o CSV. Se omitido, usa relatorios/aem-metadata/<arquivo.csv>.",
    )
    parser.add_argument(
        "--csv-encoding",
        default="utf-8-sig",
        choices=("utf-8-sig", "utf-8", "cp1252"),
        help="Encoding do CSV gerado. Padrao: utf-8-sig para importacao correta no AEM.",
    )
    parser.add_argument(
        "--excel-safe-identifiers",
        action=argparse.BooleanOptionalAction,
        default=False,
        help='Grava cdISA/cdRede como ="valor" para impedir conversao para notacao cientifica no Excel. Padrao: desabilitado para o AEM importar o valor limpo.',
    )
    parser.add_argument(
        "--normalize-aem-folders",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Normaliza pastas do assetPath como o AEM: minusculo, sem acento e espaco como hifen. Padrao: habilitado.",
    )
    parser.add_argument(
        "--verify-blob-exists",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Inclui apenas linhas cujo TargetBlobName ainda existe no container de destino. Padrao: habilitado.",
    )
    parser.add_argument(
        "--existing-blobs-json",
        help="Inventario JSON do Blob ja gerado. Quando informado, usa este arquivo para validar existencia em vez de listar o container.",
    )
    args = parser.parse_args()
    blob_prefixes = normalize_blob_prefixes(args.blob_prefix)

    output_dir = Path(args.output_dir).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    settings = get_function_app_settings(args.resource_group, args.function_app_name)
    table_name = args.table_name or settings.get("STATUS_TABLE_NAME")
    if not table_name:
        raise SystemExit("STATUS_TABLE_NAME nao informado e nao encontrado na Function App.")

    connection_string = get_connection_string(args, settings)
    rows = load_status_rows_for_prefixes(
        table_name,
        connection_string,
        args.status,
        blob_prefixes,
        args.include_empty_metadata,
        args.max_rows,
    )

    existing_blobs = None
    if args.verify_blob_exists:
        if args.existing_blobs_json:
            existing_blobs = load_existing_blob_names_from_inventory(args.existing_blobs_json, blob_prefixes)
        else:
            existing_blobs = load_existing_blob_names(args, settings, connection_string, blob_prefixes)

    csv_rows, skipped = build_aem_rows(
        rows,
        args.aem_root,
        args.strip_blob_prefix,
        args.normalize_aem_folders,
        existing_blobs,
    )

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_prefix = slug_path(single_output_slug(blob_prefixes))
    csv_path = output_dir / f"aem_bulk_metadata_{safe_prefix}_{timestamp}.csv"
    summary_path = output_dir / f"aem_bulk_metadata_{safe_prefix}_{timestamp}_resumo.json"

    write_aem_csv(csv_path, csv_rows, args.csv_encoding, args.excel_safe_identifiers)
    uploaded_blob = ""
    if args.upload_to_blob:
        uploaded_blob = upload_csv_to_blob(csv_path, args, settings, connection_string)

    summary = {
        "generatedAt": datetime.now().isoformat(timespec="seconds"),
        "tableName": table_name,
        "status": args.status,
        "blobPrefixes": blob_prefixes,
        "stripBlobPrefix": args.strip_blob_prefix,
        "aemRoot": args.aem_root,
        "rowsRead": len(rows),
        "rowsWritten": len(csv_rows),
        "rowsSkipped": skipped,
        "csvEncoding": args.csv_encoding,
        "excelSafeIdentifiers": args.excel_safe_identifiers,
        "normalizeAemFolders": args.normalize_aem_folders,
        "verifyBlobExists": args.verify_blob_exists,
        "existingBlobsRead": len(existing_blobs) if existing_blobs is not None else None,
        "csv": str(csv_path),
        "uploadedBlob": uploaded_blob,
    }
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    print("CONCLUIDO")
    print(f"Tabela: {table_name}")
    print(f"Linhas lidas: {len(rows)}")
    print(f"Linhas gravadas: {len(csv_rows)}")
    print(f"Ignoradas: {skipped}")
    print(f"CSV: {csv_path}")
    if uploaded_blob:
        print(f"Blob CSV: {uploaded_blob}")
    print(f"Resumo: {summary_path}")
    return 0


def get_function_app_settings(resource_group: str, function_app_name: str) -> dict[str, str]:
    az_cli = find_azure_cli()
    command = [
        az_cli,
        "functionapp",
        "config",
        "appsettings",
        "list",
        "-g",
        resource_group,
        "-n",
        function_app_name,
        "-o",
        "json",
    ]
    try:
        result = subprocess.run(command, capture_output=True, text=True, check=True)
    except subprocess.CalledProcessError as exc:
        raise RuntimeError(f"Falha ao ler app settings da Function: {exc.stderr.strip()}") from exc

    values = json.loads(result.stdout)
    return {item["name"]: item.get("value", "") for item in values}


def get_connection_string(args: argparse.Namespace, settings: dict[str, str]) -> str:
    connection_string = (
        args.connection_string
        or os.environ.get("STATUS_STORAGE_CONNECTION_STRING")
        or settings.get("STATUS_STORAGE_CONNECTION_STRING")
        or os.environ.get("AzureWebJobsStorage")
        or settings.get("AzureWebJobsStorage")
    )
    if not connection_string:
        raise SystemExit("Connection string nao encontrada. Informe --connection-string.")
    return connection_string


def normalize_blob_prefixes(values: list[str]) -> list[str]:
    prefixes = []
    seen = set()
    for value in values:
        prefix = str(value or "").strip("/")
        if prefix and prefix not in seen:
            prefixes.append(prefix)
            seen.add(prefix)
    return prefixes


def single_output_slug(blob_prefixes: list[str]) -> str:
    if not blob_prefixes:
        return "todos"
    if len(blob_prefixes) == 1:
        return blob_prefixes[0]
    common_parts: list[str] = []
    split_prefixes = [prefix.split("/") for prefix in blob_prefixes]
    for parts in zip(*split_prefixes):
        if len(set(parts)) != 1:
            break
        common_parts.append(parts[0])
    if common_parts:
        return "/".join(common_parts) + "_consolidado"
    return "consolidado"


def load_status_rows_for_prefixes(
    table_name: str,
    connection_string: str,
    status: str,
    blob_prefixes: list[str],
    include_empty_metadata: bool,
    max_rows: int | None,
) -> list[dict[str, Any]]:
    if not blob_prefixes:
        return load_status_rows(table_name, connection_string, status, "", include_empty_metadata, max_rows)

    rows: list[dict[str, Any]] = []
    seen_blob_names: set[str] = set()
    for blob_prefix in blob_prefixes:
        remaining = None if max_rows is None else max_rows - len(rows)
        if remaining is not None and remaining <= 0:
            break
        for row in load_status_rows(table_name, connection_string, status, blob_prefix, include_empty_metadata, remaining):
            blob_name = str(row.get("TargetBlobName") or "").strip("/")
            if blob_name in seen_blob_names:
                continue
            rows.append(row)
            seen_blob_names.add(blob_name)
            if max_rows is not None and len(rows) >= max_rows:
                return rows
    return rows


def load_status_rows(
    table_name: str,
    connection_string: str,
    status: str,
    blob_prefix: str,
    include_empty_metadata: bool,
    max_rows: int | None,
) -> list[dict[str, Any]]:
    filters = ["PartitionKey eq 'migration'", f"Status eq '{escape_odata(status)}'"]
    if blob_prefix:
        filters.append(f"TargetBlobName ge '{escape_odata(blob_prefix)}'")
    query_filter = " and ".join(filters)

    az_cli = find_azure_cli()
    rows: list[dict[str, Any]] = []
    marker: dict[str, str] | None = None

    while True:
        command = [
            az_cli,
            "storage",
            "entity",
            "query",
            "--table-name",
            table_name,
            "--filter",
            query_filter,
            "--connection-string",
            connection_string,
            "--num-results",
            "1000",
            "-o",
            "json",
        ]
        if marker:
            command.extend(["--marker", f"nextpartitionkey={marker['nextpartitionkey']}", f"nextrowkey={marker['nextrowkey']}"])

        try:
            result = subprocess.run(command, capture_output=True, text=True, check=True)
        except subprocess.CalledProcessError as exc:
            raise RuntimeError(f"Falha ao consultar tabela {table_name}: {exc.stderr.strip()}") from exc

        payload = json.loads(result.stdout or "{}")
        for entity in payload.get("items") or []:
            blob_name = str(entity.get("TargetBlobName") or "")
            if not blob_name:
                continue
            if blob_prefix and not blob_name.startswith(blob_prefix):
                continue
            if not include_empty_metadata and not entity.get("AemMetadataJson"):
                continue
            rows.append(dict(entity))
            if max_rows is not None and len(rows) >= max_rows:
                return rows

        next_marker = payload.get("nextMarker") or payload.get("next_marker") or {}
        next_partition = next_marker.get("nextpartitionkey") or next_marker.get("nextPartitionKey")
        next_row = next_marker.get("nextrowkey") or next_marker.get("nextRowKey")
        if not next_partition or not next_row:
            break
        marker = {"nextpartitionkey": next_partition, "nextrowkey": next_row}

    return rows


def build_aem_rows(
    rows: list[dict[str, Any]],
    aem_root: str,
    strip_blob_prefix: str,
    normalize_aem_folders: bool = True,
    existing_blobs: set[str] | None = None,
) -> tuple[list[dict[str, str]], int]:
    output: list[dict[str, str]] = []
    skipped = 0
    seen_asset_paths: set[str] = set()

    for row in rows:
        blob_name = str(row.get("TargetBlobName") or "").strip("/")
        if not blob_name:
            skipped += 1
            continue
        if is_zero_byte_status_row(row):
            skipped += 1
            continue
        if is_excluded_aem_metadata_blob(blob_name):
            skipped += 1
            continue
        if existing_blobs is not None and blob_name not in existing_blobs:
            skipped += 1
            continue

        metadata = parse_metadata(row.get("AemMetadataJson"))
        if not metadata:
            skipped += 1
            continue
        source_name = str(row.get("SourceName") or "").strip()
        if source_name:
            metadata.setdefault("demo:nomeArquivoOriginal", source_name)
        metadata.update(dam_metadata_from_blob_name(blob_name))
        asset_path = build_asset_path(aem_root, blob_name, strip_blob_prefix, normalize_aem_folders)
        if asset_path in seen_asset_paths:
            skipped += 1
            continue
        seen_asset_paths.add(asset_path)

        item = {"assetPath": asset_path}
        for column, _type_name in DEFAULT_METADATA_COLUMNS:
            item[column] = clean_csv_value(metadata.get(column, ""), column)
        output.append(item)

    output.sort(key=lambda item: item["assetPath"].lower())
    return output, skipped


def is_excluded_aem_metadata_blob(blob_name: str) -> bool:
    return PurePosixPath(str(blob_name)).suffix.lower() in EXCLUDED_AEM_METADATA_EXTENSIONS


def is_zero_byte_status_row(row: dict[str, Any]) -> bool:
    size = row.get("SizeBytes")
    if size in (None, ""):
        return False
    try:
        return int(size) == 0
    except (TypeError, ValueError):
        return False


def dam_metadata_from_blob_name(blob_name: str) -> dict[str, str]:
    metadata = {"demo:caminhoBlobNormalizado": blob_name}
    parts = [part for part in blob_name.strip("/").split("/") if part]
    if len(parts) >= 7 and parts[0] == "sp" and parts[1] == "ativos":
        categoria = ""
        origem_index = 6
        rule = "sp/ativos/{ciclo}/{tipo}/{codigoColecao}/{retrancaMaterial}/{origem}"
        if len(parts) >= 8 and parts[6].lower() in DAM_CATEGORY_BY_SEGMENT:
            categoria = DAM_CATEGORY_BY_SEGMENT[parts[6].lower()]
            origem_index = 7
            rule = "sp/ativos/{ciclo}/{tipo}/{codigoColecao}/{retrancaMaterial}/{categoria}/{origem}"

        metadata.update(
            {
                "demo:damRoot": "sp/ativos",
                "demo:damCiclo": parts[2],
                "demo:damTipo": parts[3],
                "demo:damCodigoColecao": parts[4],
                "demo:damRetrancaMaterial": parts[5],
                "demo:damCategoria": categoria,
                "demo:damOrigemConteudo": friendly_origin(parts[origem_index]),
                "demo:damRegraTaxonomia": rule,
            }
        )
    return metadata


def friendly_origin(value: str) -> str:
    normalized = value.strip().lower()
    if normalized == "arquivos-abertos":
        return "Arquivos Abertos"
    if normalized == "pdf-final":
        return "PDF - FINAL"
    return value


def load_existing_blob_names(
    args: argparse.Namespace,
    settings: dict[str, str],
    connection_string: str,
    blob_prefixes: list[str],
) -> set[str]:
    container_name = settings.get("DEST_CONTAINER_NAME")
    if not container_name:
        raise SystemExit("DEST_CONTAINER_NAME nao encontrado na Function App.")

    az_cli = find_azure_cli()
    prefixes = blob_prefixes or [""]
    names: set[str] = set()
    for blob_prefix in prefixes:
        command = [
            az_cli,
            "storage",
            "blob",
            "list",
            "--container-name",
            container_name,
            "--prefix",
            blob_prefix,
            "--connection-string",
            connection_string,
            "-o",
            "json",
        ]
        try:
            result = subprocess.run(command, capture_output=True, text=True, check=True)
        except subprocess.CalledProcessError as exc:
            raise RuntimeError(f"Falha ao listar blobs do container {container_name}: {exc.stderr.strip()}") from exc

        payload = json.loads(result.stdout or "[]")
        names.update(str(item.get("name") or "").strip("/") for item in payload if item.get("name"))
    return names


def load_existing_blob_names_from_inventory(path_value: str, blob_prefixes: list[str]) -> set[str]:
    inventory_path = Path(path_value).expanduser().resolve()
    payload = json.loads(inventory_path.read_text(encoding="utf-8-sig"))
    items = payload.get("items", []) if isinstance(payload, dict) else payload
    if not isinstance(items, list):
        raise SystemExit(f"Inventario Blob invalido: {inventory_path}")

    prefixes = [prefix.strip("/") for prefix in blob_prefixes if prefix.strip("/")]
    names: set[str] = set()
    for item in items:
        if not isinstance(item, dict):
            continue
        name = str(item.get("path") or item.get("name") or "").strip("/")
        if not name:
            continue
        size = item.get("size")
        if size not in (None, ""):
            try:
                if int(size) == 0:
                    continue
            except (TypeError, ValueError):
                pass
        if prefixes and not any(name == prefix or name.startswith(prefix + "/") for prefix in prefixes):
            continue
        names.add(name)
    return names


def parse_metadata(raw: object) -> dict[str, str]:
    if not raw:
        return {}
    try:
        values = json.loads(str(raw))
    except json.JSONDecodeError:
        return {}
    if not isinstance(values, dict):
        return {}
    return {str(key): "" if value is None else str(value) for key, value in values.items()}


def build_asset_path(
    aem_root: str,
    blob_name: str,
    strip_blob_prefix: str,
    normalize_aem_folders: bool = True,
) -> str:
    relative = blob_name.strip("/")
    prefix = strip_blob_prefix.strip("/")
    if prefix and relative.lower().startswith(prefix.lower() + "/"):
        relative = relative[len(prefix) + 1 :]
    elif prefix and relative.lower() == prefix.lower():
        relative = ""

    if normalize_aem_folders:
        relative = normalize_aem_relative_path(relative)

    joined = "/".join(part.strip("/") for part in [aem_root, relative] if part.strip("/"))
    return f"/{joined}" if joined else "/"


def normalize_aem_relative_path(relative: str) -> str:
    parts = [part for part in relative.strip("/").split("/") if part]
    if len(parts) <= 1:
        return "/".join(parts)

    folders = [normalize_aem_folder_segment(part) for part in parts[:-1]]
    return "/".join([*folders, parts[-1]])


def normalize_aem_folder_segment(value: str) -> str:
    text = remove_accents(value).lower()
    text = re.sub(r"\s*-\s*", "-", text)
    text = re.sub(r"[\s_.]+", "-", text)
    text = re.sub(r"[^a-z0-9_-]", "-", text)
    text = re.sub(r"-{2,}", "-", text)
    return text.strip("-.") or "pasta"


def remove_accents(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    return "".join(char for char in normalized if not unicodedata.combining(char))


def write_aem_csv(
    path: Path,
    rows: list[dict[str, str]],
    encoding: str = "utf-8-sig",
    excel_safe_identifiers: bool = False,
) -> None:
    columns = ["assetPath", *[column for column, _type_name in DEFAULT_METADATA_COLUMNS]]
    header = ["assetPath", *[f"{column}{{{{{type_name}}}}}" for column, type_name in DEFAULT_METADATA_COLUMNS]]
    with path.open("w", encoding=encoding, newline="") as fp:
        fp.write("sep=,\n")
        writer = csv.writer(fp, delimiter=",", lineterminator="\n", quoting=csv.QUOTE_ALL)
        writer.writerow(header)
        for row in rows:
            writer.writerow([format_output_value(column, row.get(column, ""), excel_safe_identifiers) for column in columns])


def format_output_value(column: str, value: str, excel_safe_identifiers: bool) -> str:
    text = str(value)
    if excel_safe_identifiers and column in IDENTIFIER_COLUMNS and text:
        escaped = text.replace('"', '""')
        return f'="{escaped}"'
    return text


def upload_csv_to_blob(
    path: Path,
    args: argparse.Namespace,
    settings: dict[str, str],
    connection_string: str,
) -> str:
    container_name = settings.get("DEST_CONTAINER_NAME")
    if not container_name:
        raise SystemExit("DEST_CONTAINER_NAME nao encontrado na Function App.")

    blob_name = args.output_blob_name or f"relatorios/aem-metadata/{path.name}"
    az_cli = find_azure_cli()
    command = [
        az_cli,
        "storage",
        "blob",
        "upload",
        "--container-name",
        container_name,
        "--name",
        blob_name,
        "--file",
        str(path),
        "--connection-string",
        connection_string,
        "--content-type",
        f"text/csv; charset={args.csv_encoding.replace('-sig', '')}",
        "--overwrite",
        "true",
        "-o",
        "none",
    ]
    try:
        subprocess.run(command, capture_output=True, text=True, check=True)
    except subprocess.CalledProcessError as exc:
        raise RuntimeError(f"Falha ao subir CSV no Blob: {exc.stderr.strip()}") from exc
    return blob_name


def clean_csv_value(value: str, column: str | None = None) -> str:
    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    text = " ".join(part.strip() for part in text.split("\n") if part.strip())
    if column in IDENTIFIER_COLUMNS:
        return normalize_identifier(text)
    if column in INTEGER_COLUMNS:
        return normalize_integer(text)
    return text


def normalize_identifier(value: str) -> str:
    text = value.strip()
    if not text:
        return ""
    candidate = text.replace(",", ".")
    if not re.fullmatch(r"[+-]?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?", candidate):
        return text
    if "e" not in candidate.lower():
        return text[:-2] if text.endswith(".0") else text
    try:
        number = Decimal(candidate)
    except InvalidOperation:
        return text
    if number == number.to_integral_value():
        return format(number.quantize(Decimal(1)), "f")
    return format(number.normalize(), "f").rstrip("0").rstrip(".")


def normalize_integer(value: str) -> str:
    text = value.strip()
    candidate = text.replace(",", ".")
    if not re.fullmatch(r"[+-]?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?", candidate):
        return text
    try:
        number = Decimal(candidate)
    except InvalidOperation:
        return text
    if number == number.to_integral_value():
        return format(number.quantize(Decimal(1)), "f")
    return text


def escape_odata(value: str) -> str:
    return value.replace("'", "''")


def slug_path(value: str) -> str:
    text = value.strip("/").replace("\\", "/")
    text = "".join(ch if ch.isalnum() else "-" for ch in text)
    text = "-".join(part for part in text.split("-") if part)
    return text[:80] or "todos"


def find_azure_cli() -> str:
    for candidate in ("az", "az.cmd", "az.exe"):
        path = shutil.which(candidate)
        if path:
            return path

    windows_default = Path(r"C:\Program Files\Microsoft SDKs\Azure\CLI2\wbin\az.cmd")
    if windows_default.exists():
        return str(windows_default)

    return "az"


if __name__ == "__main__":
    sys.exit(main())

