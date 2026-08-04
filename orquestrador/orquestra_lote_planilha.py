#!/usr/bin/env python
from __future__ import annotations

import argparse
import base64
import json
import shutil
import subprocess
import sys
import time
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any


REPO_ROOT = Path(__file__).resolve().parents[1]
MENSAGEM_SCRIPT = REPO_ROOT / "montar-fila" / "monta_lote_planilha_sp.py"

def main() -> int:
    parser = argparse.ArgumentParser(
        description="Orquestra leitura da planilha e envio para fila. O CSV AEM final deve ser gerado em script separado apos validacao."
    )
    parser.add_argument("--workbook", required=True, help="Caminho da planilha .xlsx.")
    parser.add_argument("--sheet", default="ATIVOS-LT-1-2", help="Nome da aba.")
    parser.add_argument("--start-row", type=int, default=5, help="Primeira linha valida.")
    parser.add_argument("--end-row", type=int, help="Ultima linha da planilha a considerar.")
    parser.add_argument("--row", type=int, action="append", help="Processa apenas esta linha. Pode repetir.")
    parser.add_argument("--max-rows", type=int, help="Limita a quantidade de linhas validas.")
    parser.add_argument(
        "--filter-tag",
        action="append",
        help="Processa somente linhas cuja coluna TAGS - FASE 1 contenha esta tag. Pode repetir.",
    )
    parser.add_argument("--target-blob-root", default="sp/ativos", help="Raiz no Blob. Padrao: sp/ativos.")
    parser.add_argument(
        "--aem-root",
        default="",
        help="Reservado para fluxos externos de metadata AEM.",
    )
    parser.add_argument(
        "--output-dir",
        default="",
        help="Pasta de saida. Se omitido, usa tmp/orquestrador/<timestamp> no workspace.",
    )
    parser.add_argument("--resource-group", default="rg-demo-dam", help="Resource group da Function.")
    parser.add_argument(
        "--function-app-name",
        default="sp-dam-migration-worker",
        help="Nome da Function App.",
    )
    parser.add_argument(
        "--send-to-queue",
        action="store_true",
        help="Envia as mensagens geradas para a fila de pastas.",
    )
    parser.add_argument(
        "--wait",
        action="store_true",
        help="Aguarda as filas esvaziarem antes de concluir a orquestracao.",
    )
    parser.add_argument(
        "--upload-metadata-csv",
        action="store_true",
        help="Nao usado pelo orquestrador. Use aem-metadata\\gera_csv_bulk_aem.py para gerar metadata AEM.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Gera as mensagens e relatorios, mas nao envia para fila nem sobe CSV.",
    )
    parser.add_argument("--poll-seconds", type=int, default=30, help="Intervalo de consulta das filas.")
    parser.add_argument("--timeout-minutes", type=int, default=240, help="Timeout da espera.")
    parser.add_argument(
        "--settle-checks",
        type=int,
        default=3,
        help="Quantidade de leituras zeradas consecutivas antes de considerar concluido.",
    )
    parser.add_argument(
        "--strip-blob-prefix",
        default="",
        help="Reservado para fluxos externos de metadata AEM.",
    )
    parser.add_argument(
        "--raw-queue-content",
        action="store_true",
        help="Envia JSON puro na fila. Padrao: envia Base64, como Storage Explorer.",
    )
    args = parser.parse_args()
    if args.upload_metadata_csv:
        raise SystemExit(
            "--upload-metadata-csv nao e processado pelo orquestrador. "
            "Use aem-metadata\\gera_csv_bulk_aem.py para gerar metadata AEM."
        )
    if args.end_row is not None and args.end_row < args.start_row:
        raise SystemExit("--end-row deve ser maior ou igual a --start-row.")

    output_dir = resolve_output_dir(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    messages_dir = output_dir / "mensagens"
    messages_dir.mkdir(parents=True, exist_ok=True)

    print("1/4 Gerando mensagens da planilha...")
    run_message_generator(args, messages_dir)
    messages = read_messages(messages_dir / "mensagens-fila.jsonl")
    if not messages:
        raise SystemExit("Nenhuma mensagem foi gerada. Confira a planilha/filtros.")
    print(f"Mensagens geradas: {len(messages)}")

    settings = get_function_app_settings(args.resource_group, args.function_app_name)
    connection_string = get_connection_string(settings)
    folder_queue = settings.get("FOLDER_QUEUE_NAME")
    file_queue = settings.get("QUEUE_NAME")
    if not folder_queue or not file_queue:
        raise SystemExit("FOLDER_QUEUE_NAME ou QUEUE_NAME nao encontrado nas app settings.")

    queue_names = {
        "folders": folder_queue,
        "files": file_queue,
        "foldersPoison": f"{folder_queue}-poison",
        "filesPoison": f"{file_queue}-poison",
    }
    start_counts = get_queue_counts(queue_names, connection_string)

    if args.dry_run:
        print("Dry-run ativo: mensagens nao foram enviadas.")
    elif args.send_to_queue:
        print("2/4 Enviando mensagens para a fila de pastas...")
        send_messages_to_queue(folder_queue, messages, connection_string, raw=args.raw_queue_content)
        print(f"Mensagens enviadas para {folder_queue}: {len(messages)}")
    else:
        print("Envio para fila desativado. Use --send-to-queue para enviar.")

    if args.wait and not args.dry_run and args.send_to_queue:
        print("3/4 Aguardando processamento das filas...")
        wait_for_processing(queue_names, connection_string, start_counts, args.poll_seconds, args.timeout_minutes, args.settle_checks)
    elif args.wait:
        print("Espera ignorada porque --send-to-queue nao foi usado ou dry-run esta ativo.")

    summary = {
        "generatedAt": datetime.now().isoformat(timespec="seconds"),
        "workbook": str(Path(args.workbook).expanduser().resolve()),
        "messages": len(messages),
        "filterTags": args.filter_tag or [],
        "sentToQueue": bool(args.send_to_queue and not args.dry_run),
        "waited": bool(args.wait and args.send_to_queue and not args.dry_run),
        "metadataCsvGenerated": False,
        "metadataCsvNote": "CSV AEM final deve ser gerado separadamente apos inventario/validacao do Blob.",
        "outputDir": str(output_dir),
    }
    summary_path = output_dir / "orquestracao-resumo.json"
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print("4/4 Concluido.")
    print(f"Resumo: {summary_path}")
    print("Metadata AEM: use aem-metadata\\gera_csv_bulk_aem.py quando necessario.")
    return 0


def resolve_output_dir(value: str) -> Path:
    if value:
        return Path(value).expanduser().resolve()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return (REPO_ROOT.parents[1] / "tmp" / "orquestrador" / timestamp).resolve()


def run_message_generator(args: argparse.Namespace, output_dir: Path) -> None:
    command = [
        sys.executable,
        str(MENSAGEM_SCRIPT),
        "--workbook",
        str(Path(args.workbook).expanduser().resolve()),
        "--sheet",
        args.sheet,
        "--start-row",
        str(args.start_row),
        "--target-blob-root",
        args.target_blob_root,
        "--output-dir",
        str(output_dir),
        "--resource-group",
        args.resource_group,
        "--function-app-name",
        args.function_app_name,
    ]
    for row in args.row or []:
        command.extend(["--row", str(row)])
    if args.end_row is not None:
        command.extend(["--end-row", str(args.end_row)])
    if args.max_rows is not None:
        command.extend(["--max-rows", str(args.max_rows)])
    for filter_tag in args.filter_tag or []:
        command.extend(["--filter-tag", filter_tag])
    run(command)


def read_messages(path: Path) -> list[dict[str, Any]]:
    messages = []
    with path.open("r", encoding="utf-8") as fp:
        for line in fp:
            line = line.strip()
            if line:
                messages.append(json.loads(line))
    return messages


def send_messages_to_queue(queue_name: str, messages: list[dict[str, Any]], connection_string: str, *, raw: bool) -> None:
    az_cli = find_azure_cli()
    for index, message in enumerate(messages, start=1):
        content = json.dumps(message, ensure_ascii=False, separators=(",", ":"))
        if not raw:
            content = base64.b64encode(content.encode("utf-8")).decode("ascii")
        run(
            [
                az_cli,
                "storage",
                "message",
                "put",
                "--queue-name",
                queue_name,
                "--content",
                content,
                "--connection-string",
                connection_string,
                "-o",
                "none",
            ]
        )
        print(f"  enviada {index}/{len(messages)}")


def wait_for_processing(
    queue_names: dict[str, str],
    connection_string: str,
    start_counts: dict[str, int],
    poll_seconds: int,
    timeout_minutes: int,
    settle_checks: int,
) -> None:
    deadline = time.time() + timeout_minutes * 60
    stable_zero = 0
    while True:
        counts = get_queue_counts(queue_names, connection_string)
        new_folder_poison = max(0, counts["foldersPoison"] - start_counts.get("foldersPoison", 0))
        new_file_poison = max(0, counts["filesPoison"] - start_counts.get("filesPoison", 0))
        print(
            "  filas: "
            f"folders={counts['folders']} files={counts['files']} "
            f"folders-poison={counts['foldersPoison']} files-poison={counts['filesPoison']}"
        )
        if new_folder_poison or new_file_poison:
            raise SystemExit(
                "Processamento gerou mensagens poison novas. "
                f"folders-poison +{new_folder_poison}, files-poison +{new_file_poison}."
            )
        if counts["folders"] == 0 and counts["files"] == 0:
            stable_zero += 1
            if stable_zero >= settle_checks:
                return
        else:
            stable_zero = 0
        if time.time() >= deadline:
            raise SystemExit("Timeout aguardando processamento das filas.")
        time.sleep(poll_seconds)


def get_queue_counts(queue_names: dict[str, str], connection_string: str) -> dict[str, int]:
    return {
        key: get_queue_count(queue_name, connection_string)
        for key, queue_name in queue_names.items()
    }


def get_queue_count(queue_name: str, connection_string: str) -> int:
    az_cli = find_azure_cli()
    result = run(
        [
            az_cli,
            "storage",
            "queue",
            "metadata",
            "show",
            "--name",
            queue_name,
            "--connection-string",
            connection_string,
            "-o",
            "json",
        ],
        capture=True,
    )
    payload = json.loads(result.stdout or "{}")
    return int(payload.get("approximateMessageCount") or 0)


def group_collection_prefixes(messages: list[dict[str, Any]]) -> dict[str, str]:
    groups: dict[str, str] = {}
    for message in messages:
        target_prefix = str(message.get("targetBlobPrefix") or "").strip("/")
        parts = target_prefix.split("/")
        if len(parts) >= 5 and parts[0] == "sp" and parts[1] == "ativos":
            collection_prefix = "/".join(parts[:5])
            groups[collection_prefix] = parts[4]
        elif len(parts) >= 2:
            collection_prefix = "/".join(parts[:2])
            groups[collection_prefix] = parts[1]
    return dict(sorted(groups.items()))


def metadata_csv_blob_root(target_blob_root: str) -> str:
    parts = [part for part in target_blob_root.strip("/").split("/") if part]
    if parts and parts[0] == "sp":
        return "sp"
    return "/".join(parts[:1]) or "sp"


def metadata_csv_run_slug(args: argparse.Namespace) -> str:
    if args.filter_tag:
        return slug_file("_".join(args.filter_tag))
    if args.end_row is not None:
        return f"linhas_{args.start_row}_{args.end_row}"
    return f"linhas_{args.start_row}_fim"


def get_function_app_settings(resource_group: str, function_app_name: str) -> dict[str, str]:
    az_cli = find_azure_cli()
    result = run(
        [
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
        ],
        capture=True,
    )
    return {item["name"]: item.get("value", "") for item in json.loads(result.stdout)}


def get_connection_string(settings: dict[str, str]) -> str:
    connection_string = settings.get("QUEUE_STORAGE_CONNECTION_STRING") or settings.get("AzureWebJobsStorage")
    if not connection_string:
        raise SystemExit("QUEUE_STORAGE_CONNECTION_STRING ou AzureWebJobsStorage nao encontrado.")
    return connection_string


def slug_file(value: str) -> str:
    text = "".join(ch if ch.isalnum() else "_" for ch in value)
    text = "_".join(part for part in text.split("_") if part)
    return text[:120] or "colecao"


def find_azure_cli() -> str:
    for candidate in ("az", "az.cmd", "az.exe"):
        path = shutil.which(candidate)
        if path:
            return path
    windows_default = Path(r"C:\Program Files\Microsoft SDKs\Azure\CLI2\wbin\az.cmd")
    if windows_default.exists():
        return str(windows_default)
    return "az"


def run(command: list[str], *, capture: bool = False) -> subprocess.CompletedProcess[str]:
    try:
        return subprocess.run(command, capture_output=capture, text=True, check=True)
    except subprocess.CalledProcessError as exc:
        stderr = exc.stderr.strip() if exc.stderr else ""
        stdout = exc.stdout.strip() if exc.stdout else ""
        detail = stderr or stdout or str(exc)
        raise RuntimeError(f"Comando falhou: {' '.join(command)}\n{detail}") from exc


if __name__ == "__main__":
    sys.exit(main())

