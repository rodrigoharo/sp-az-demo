import re
import unicodedata
from urllib.parse import urlparse

from azure.core.exceptions import ResourceExistsError
from azure.data.tables import TableClient, UpdateMode
from azure.identity import DefaultAzureCredential
from azure.storage.blob import BlobClient, ContainerClient, ContentSettings
from azure.storage.queue import QueueClient, TextBase64EncodePolicy

from .config import optional_setting, require_setting


def _credential() -> DefaultAzureCredential:
    return DefaultAzureCredential()


def get_container_client() -> ContainerClient:
    account_url = require_setting("DEST_STORAGE_ACCOUNT_URL")
    container_name = require_setting("DEST_CONTAINER_NAME")
    return ContainerClient(account_url=account_url, container_name=container_name, credential=_credential())


def get_blob_client(blob_name: str) -> BlobClient:
    return get_container_client().get_blob_client(blob_name)


def upload_blob(
    blob_name: str,
    content: object,
    *,
    content_type: str | None = None,
    metadata: dict[str, str] | None = None,
    tags: dict[str, str] | None = None,
    length: int | None = None,
) -> None:
    blob = get_blob_client(blob_name)
    settings = ContentSettings(content_type=content_type) if content_type else None
    kwargs = {
        "overwrite": True,
        "content_settings": settings,
        "metadata": _clean_dict(metadata),
        "tags": _clean_dict(tags),
    }
    if length is not None:
        kwargs["length"] = length

    blob.upload_blob(content, **kwargs)


def get_status_table_client() -> TableClient:
    table_name = require_setting("STATUS_TABLE_NAME")
    connection_string = optional_setting("STATUS_STORAGE_CONNECTION_STRING")
    if connection_string:
        return TableClient.from_connection_string(connection_string, table_name=table_name)

    account_url = require_setting("DEST_STORAGE_ACCOUNT_URL")
    storage_account = urlparse(account_url).netloc.split(".", 1)[0]
    table_endpoint = f"https://{storage_account}.table.core.windows.net"
    return TableClient(endpoint=table_endpoint, table_name=table_name, credential=_credential())


def ensure_status_table() -> None:
    try:
        get_status_table_client().create_table()
    except ResourceExistsError:
        return


def upsert_status(entity: dict[str, object]) -> None:
    table = get_status_table_client()
    table.upsert_entity(entity=entity, mode=UpdateMode.REPLACE)


def get_queue_client(queue_name: str) -> QueueClient:
    connection_string = require_setting("QUEUE_STORAGE_CONNECTION_STRING")
    return QueueClient.from_connection_string(
        connection_string,
        queue_name=queue_name,
        message_encode_policy=TextBase64EncodePolicy(),
    )


def _clean_dict(values: dict[str, object] | None) -> dict[str, str] | None:
    if not values:
        return None
    cleaned = {}
    for key, value in values.items():
        if value is None:
            continue
        clean_key = _clean_metadata_key(str(key))
        clean_value = _clean_metadata_value(str(value))
        if clean_key and clean_value:
            cleaned[clean_key] = clean_value
    return cleaned or None


def _clean_metadata_key(value: str) -> str:
    clean = re.sub(r"[^A-Za-z0-9_]", "_", value.strip())
    if clean and clean[0].isdigit():
        clean = f"m_{clean}"
    return clean[:128]


def _clean_metadata_value(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    ascii_value = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    ascii_value = ascii_value.encode("ascii", "ignore").decode("ascii")
    ascii_value = re.sub(r"[\r\n\t]+", " ", ascii_value)
    ascii_value = re.sub(r"\s+", " ", ascii_value).strip()
    return ascii_value[:1024]

