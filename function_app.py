import json
import logging
import mimetypes
from datetime import datetime, timezone
from pathlib import PurePosixPath
from uuid import uuid4

import azure.functions as func

from shared_code.config import require_setting
from shared_code.dam_normalization import normalize_dam_relative_path
from shared_code.graph_client import GraphClient
from shared_code.storage_clients import ensure_status_table, get_queue_client, upload_blob, upsert_status


app = func.FunctionApp()


MIME_TYPE_OVERRIDES = {
    ".psb": "image/vnd.adobe.photoshop",
}

EXCLUDED_TEMPORARY_EXTENSIONS = {
    ".bmap",
    ".idlk",
    ".lst 2",
    ".textclipping",
}


@app.queue_trigger(
    arg_name="msg",
    queue_name="%QUEUE_NAME%",
    connection="QUEUE_STORAGE_CONNECTION_STRING",
)
def migrate_file_worker(msg: func.QueueMessage) -> None:
    raw = msg.get_body().decode("utf-8")
    message_id = msg.id or str(uuid4())
    logging.debug("Processing queue message %s", message_id)

    entity_base = {
        "PartitionKey": "migration",
        "RowKey": message_id,
        "QueueMessageId": message_id,
        "StartedAtUtc": _utc_now(),
    }

    message: dict = {}
    source_file = None

    try:
        message = json.loads(raw)
        target_blob_name = message.get("targetBlobName")
        if not target_blob_name:
            raise ValueError("Queue message must contain targetBlobName.")
        target_blob_name = normalize_dam_relative_path(str(target_blob_name).strip("/"))

        if message.get("sharePointSharingUrl") or message.get("sharePointDrivePath"):
            source_file = GraphClient().open_stream_from_message(message)
            content = source_file.content
            source_name = source_file.name
            source_url = source_file.source_url
            source_status = _source_status_from_graph_file(source_file)
            size_bytes = source_file.size
            source_mime_type = source_file.mime_type
        else:
            content = _test_content(message_id)
            source_name = "generated-test.txt"
            source_url = "generated://queue-test"
            source_status = {"SourceType": "GeneratedTest"}
            size_bytes = len(content)
            source_mime_type = "text/plain"

        metadata = _blob_metadata_with_source_name(
            message.get("metadata"),
            source_name,
            target_blob_name,
            source_status,
        )
        aem_metadata = _aem_metadata_with_source_name(
            message.get("aemMetadata"),
            source_name,
            target_blob_name,
            source_status,
        )

        tags = message.get("tags")

        content_type = _content_type_for_file(source_name, message.get("contentType") or source_mime_type)
        if not _is_migratable_file(source_name, content_type):
            if source_file is not None:
                source_file.close()
            _upsert_skipped_status(
                entity_base,
                message,
                source_name,
                source_url,
                source_status,
                size_bytes,
                content_type,
            )
            return
        if _is_zero_byte_file(size_bytes):
            if source_file is not None:
                source_file.close()
            _upsert_skipped_status(
                entity_base,
                message,
                source_name,
                source_url,
                source_status,
                size_bytes,
                content_type,
                "File has 0 bytes.",
            )
            return

        try:
            upload_blob(
                target_blob_name,
                content,
                content_type=content_type,
                metadata=metadata,
                tags=tags,
                length=size_bytes,
            )
        finally:
            if source_file is not None:
                source_file.close()

        ensure_status_table()
        status_entity = {
            **entity_base,
            "Status": "Succeeded",
            "TargetBlobName": target_blob_name,
            "SourceName": source_name,
            "SourceUrl": source_url,
            "CompletedAtUtc": _utc_now(),
            **source_status,
        }
        if content_type:
            status_entity["ContentType"] = content_type
        if size_bytes is not None:
            status_entity["SizeBytes"] = size_bytes
        if aem_metadata:
            status_entity["AemMetadataJson"] = json.dumps(aem_metadata, ensure_ascii=False)[:32000]
        upsert_status(status_entity)
        logging.info(
            "Migration succeeded: messageId=%s targetBlob=%s sizeBytes=%s",
            message_id,
            target_blob_name,
            size_bytes,
        )
    except Exception as exc:
        logging.exception("Migration failed for queue message %s", message_id)
        try:
            ensure_status_table()
            upsert_status(
                {
                    **entity_base,
                    **_message_status_from_queue(message),
                    "Status": "Failed",
                    "Error": str(exc)[:32000],
                    "CompletedAtUtc": _utc_now(),
                }
            )
        finally:
            raise


@app.queue_trigger(
    arg_name="msg",
    queue_name="%FOLDER_QUEUE_NAME%",
    connection="QUEUE_STORAGE_CONNECTION_STRING",
)
def discover_sharepoint_folder(msg: func.QueueMessage) -> None:
    raw = msg.get_body().decode("utf-8")
    message_id = msg.id or str(uuid4())
    logging.debug("Processing folder discovery message %s", message_id)

    entity_base = {
        "PartitionKey": "folder-discovery",
        "RowKey": message_id,
        "QueueMessageId": message_id,
        "StartedAtUtc": _utc_now(),
    }

    try:
        message = json.loads(raw)
        hostname = message["sharePointHostname"]
        site_path = message["sharePointSitePath"]
        folder_path = message["sharePointFolderPath"].strip("/")
        drive_id = message.get("sharePointDriveId")
        target_blob_prefix = (message.get("targetBlobPrefix") or folder_path).strip("/")
        recursive = bool(message.get("recursive", True))
        metadata = message.get("metadata") or {}
        aem_metadata = message.get("aemMetadata") or {}
        tags = message.get("tags")

        file_queue = get_queue_client(require_setting("QUEUE_NAME"))
        graph = GraphClient()
        file_count = 0
        skipped_count = 0
        used_target_blob_names: dict[str, int] = {}

        for item in graph.iter_files_from_folder(
            hostname,
            site_path,
            folder_path,
            recursive=recursive,
            drive_id=drive_id,
        ):
            source_name = item.get("name", "")
            size_bytes = item.get("size")
            content_type = _content_type_for_file(source_name, item.get("contentType"))
            if not _is_migratable_file(source_name, content_type, item.get("hasExtension")):
                skipped_count += 1
                continue
            if _is_zero_byte_file(size_bytes):
                skipped_count += 1
                continue

            drive_path = item["sharePointDrivePath"]
            item_drive_id = item.get("sharePointDriveId") or drive_id
            target_blob_name = _deduplicate_target_blob_name(
                _target_blob_name(target_blob_prefix, folder_path, drive_path),
                used_target_blob_names,
            )
            file_message = {
                "sharePointHostname": hostname,
                "sharePointSitePath": site_path,
                "sharePointDrivePath": drive_path,
                "targetBlobName": target_blob_name,
                "metadata": metadata,
            }
            if content_type:
                file_message["contentType"] = content_type
            if item_drive_id:
                file_message["sharePointDriveId"] = item_drive_id
            if aem_metadata:
                file_message["aemMetadata"] = aem_metadata
            if tags:
                file_message["tags"] = tags

            file_queue.send_message(json.dumps(file_message, ensure_ascii=False))
            file_count += 1

        ensure_status_table()
        upsert_status(
            {
                **entity_base,
                "Status": "Succeeded",
                "SourceType": "SharePointFolder",
                "SharePointHostname": hostname,
                "SharePointSitePath": site_path,
                "SharePointDrivePath": folder_path,
                "SharePointDriveId": drive_id,
                "TargetBlobName": target_blob_prefix,
                "DiscoveredFileCount": file_count,
                "SkippedFileCount": skipped_count,
                "Recursive": recursive,
                "CompletedAtUtc": _utc_now(),
            }
        )
        logging.info(
            "Folder discovery succeeded: messageId=%s folder=%s filesQueued=%s filesSkipped=%s",
            message_id,
            folder_path,
            file_count,
            skipped_count,
        )
    except Exception as exc:
        logging.exception("Folder discovery failed for queue message %s", message_id)
        try:
            ensure_status_table()
            upsert_status(
                {
                    **entity_base,
                    "Status": "Failed",
                    "SourceType": "SharePointFolder",
                    "Error": str(exc)[:32000],
                    "CompletedAtUtc": _utc_now(),
                }
            )
        finally:
            raise


def _upsert_skipped_status(
    entity_base: dict[str, object],
    message: dict,
    source_name: str,
    source_url: str,
    source_status: dict[str, object],
    size_bytes: int | None,
    content_type: str | None,
    skip_reason: str = "File has no extension or MIME type.",
) -> None:
    ensure_status_table()
    status_entity = {
        **entity_base,
        **_message_status_from_queue(message),
        "Status": "Skipped",
        "SkipReason": skip_reason,
        "SourceName": source_name,
        "SourceUrl": source_url,
        "CompletedAtUtc": _utc_now(),
        **source_status,
    }
    if content_type:
        status_entity["ContentType"] = content_type
    if size_bytes is not None:
        status_entity["SizeBytes"] = size_bytes
    aem_metadata = _aem_metadata_with_source_name(
        message.get("aemMetadata"),
        source_name,
        str(message.get("targetBlobName") or ""),
        source_status,
    )
    if aem_metadata:
        status_entity["AemMetadataJson"] = json.dumps(aem_metadata, ensure_ascii=False)[:32000]
    upsert_status(status_entity)
    logging.info(
        "Migration skipped: targetBlob=%s sourceName=%s contentType=%s",
        message.get("targetBlobName"),
        source_name,
        content_type,
    )


def _blob_metadata_with_source_name(
    raw_metadata: object,
    source_name: str,
    target_blob_name: str,
    source_status: dict[str, object],
) -> dict[str, object]:
    metadata = dict(raw_metadata or {}) if isinstance(raw_metadata, dict) else {}
    metadata.setdefault("source", "sharepoint")
    metadata["sourceName"] = source_name
    metadata["originalFileName"] = source_name
    metadata["normalizedBlobPath"] = target_blob_name
    technical_folder = _technical_folder_from_blob_name(target_blob_name)
    if technical_folder:
        metadata["technicalFolder"] = technical_folder
    original_path = _original_sharepoint_path(source_status)
    if original_path:
        metadata["sharePointOriginalPath"] = original_path
    return metadata


def _aem_metadata_with_source_name(
    raw_metadata: object,
    source_name: str,
    target_blob_name: str,
    source_status: dict[str, object],
) -> dict[str, object]:
    metadata = dict(raw_metadata or {}) if isinstance(raw_metadata, dict) else {}
    metadata["demo:nomeArquivoOriginal"] = source_name
    metadata["demo:caminhoBlobNormalizado"] = target_blob_name
    technical_folder = _technical_folder_from_blob_name(target_blob_name)
    if technical_folder:
        metadata["demo:pastaTecnica"] = technical_folder
    original_path = _original_sharepoint_path(source_status)
    if original_path:
        metadata["demo:caminhoOriginalSharePoint"] = original_path
    return metadata


def _technical_folder_from_blob_name(target_blob_name: str) -> str | None:
    parts = [part.lower() for part in str(target_blob_name).strip("/").split("/") if part]
    for protected_folder in ("links", "document-fonts"):
        if protected_folder in parts[:-1]:
            return protected_folder
    return "raiz"


def _original_sharepoint_path(source_status: dict[str, object]) -> str | None:
    drive_path = source_status.get("SharePointDrivePath")
    return str(drive_path) if drive_path else None


def _is_migratable_file(
    name: str,
    content_type: str | None,
    has_extension: bool | None = None,
) -> bool:
    if _is_ignored_system_file(name):
        return False
    if has_extension is None:
        has_extension = bool(PurePosixPath(name).suffix)
    if not has_extension:
        return False
    if PurePosixPath(str(name)).suffix.lower() in EXCLUDED_TEMPORARY_EXTENSIONS:
        return False
    if not content_type:
        return False
    return True


def _is_zero_byte_file(size_bytes: object) -> bool:
    try:
        return int(size_bytes) == 0
    except (TypeError, ValueError):
        return False


def _content_type_for_file(name: str, content_type: object = None) -> str | None:
    extension = PurePosixPath(str(name)).suffix.lower()
    if extension in MIME_TYPE_OVERRIDES:
        return MIME_TYPE_OVERRIDES[extension]
    if isinstance(content_type, str) and content_type.strip():
        return content_type.strip()
    return mimetypes.guess_type(str(name))[0]


def _is_ignored_system_file(name: str) -> bool:
    path = PurePosixPath(str(name).replace("\\", "/"))
    parts = [part.lower() for part in path.parts if part not in {"", "."}]
    if "__macosx" in parts:
        return True
    filename = parts[-1] if parts else ""
    if filename.startswith("._"):
        return True
    return filename in {".ds_store", "thumbs.db", "desktop.ini"}


def _source_status_from_graph_file(source_file) -> dict[str, object]:
    values = {
        "SourceType": source_file.source_type,
        "GraphClientId": source_file.graph_client_id,
        "SharePointHostname": source_file.sharepoint_hostname,
        "SharePointSitePath": source_file.sharepoint_site_path,
        "SharePointDrivePath": source_file.sharepoint_drive_path,
        "SharePointSiteId": source_file.sharepoint_site_id,
        "SharePointDriveId": source_file.sharepoint_drive_id,
        "SharePointItemId": source_file.sharepoint_item_id,
    }
    return {key: value for key, value in values.items() if value}


def _message_status_from_queue(message: dict) -> dict[str, object]:
    keys = (
        "targetBlobName",
        "sharePointHostname",
        "sharePointSitePath",
        "sharePointDrivePath",
        "sharePointDriveId",
        "sharePointSharingUrl",
    )
    values = {key: message.get(key) for key in keys if message.get(key)}
    return {
        _status_key_from_message_key(key): value
        for key, value in values.items()
        if value is not None
    }


def _status_key_from_message_key(key: str) -> str:
    mapping = {
        "targetBlobName": "TargetBlobName",
        "sharePointHostname": "SharePointHostname",
        "sharePointSitePath": "SharePointSitePath",
        "sharePointDrivePath": "SharePointDrivePath",
        "sharePointDriveId": "SharePointDriveId",
        "sharePointSharingUrl": "SharePointSharingUrl",
    }
    return mapping[key]


def _target_blob_name(target_blob_prefix: str, folder_path: str, drive_path: str) -> str:
    clean_folder = folder_path.strip("/")
    clean_drive_path = drive_path.strip("/")
    relative_path = clean_drive_path
    if clean_drive_path.lower().startswith(f"{clean_folder.lower()}/"):
        relative_path = clean_drive_path[len(clean_folder) + 1 :]
    elif clean_drive_path.lower() == clean_folder.lower():
        relative_path = clean_drive_path.rsplit("/", 1)[-1]

    return f"{target_blob_prefix.rstrip('/')}/{normalize_dam_relative_path(relative_path)}"


def _deduplicate_target_blob_name(target_blob_name: str, used_names: dict[str, int]) -> str:
    clean_name = target_blob_name.strip("/")
    key = clean_name.lower()
    occurrence = used_names.get(key, 0) + 1
    used_names[key] = occurrence
    if occurrence == 1:
        return clean_name

    path = PurePosixPath(clean_name)
    suffix = path.suffix
    stem = path.name[: -len(suffix)] if suffix else path.name
    unique_filename = f"{stem}-{occurrence}{suffix}"
    if str(path.parent) in {"", "."}:
        return unique_filename
    return f"{path.parent.as_posix()}/{unique_filename}"


def _test_content(message_id: str) -> bytes:
    return f"Teste Azure Function SharePoint -> Blob\nmessageId={message_id}\n".encode("utf-8")


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()

