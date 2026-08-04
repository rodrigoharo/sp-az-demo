import base64
from dataclasses import dataclass
from pathlib import PurePosixPath
import time
from typing import BinaryIO, Callable, Iterator
from urllib.parse import quote, urlparse

import requests

from .config import require_setting


GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
RETRY_STATUS_CODES = {429, 503, 504}
MAX_GRAPH_RETRIES = 6
DEFAULT_RETRY_SECONDS = 5
MAX_RETRY_SECONDS = 60


@dataclass(frozen=True)
class GraphFile:
    name: str
    content: bytes
    source_url: str
    source_type: str
    graph_client_id: str
    mime_type: str | None = None
    sharepoint_hostname: str | None = None
    sharepoint_site_path: str | None = None
    sharepoint_drive_path: str | None = None
    sharepoint_site_id: str | None = None
    sharepoint_drive_id: str | None = None
    sharepoint_item_id: str | None = None


@dataclass(frozen=True)
class GraphFileStream:
    name: str
    content: BinaryIO
    source_url: str
    source_type: str
    graph_client_id: str
    close: Callable[[], None]
    size: int | None = None
    mime_type: str | None = None
    sharepoint_hostname: str | None = None
    sharepoint_site_path: str | None = None
    sharepoint_drive_path: str | None = None
    sharepoint_site_id: str | None = None
    sharepoint_drive_id: str | None = None
    sharepoint_item_id: str | None = None


class GraphClient:
    def __init__(self) -> None:
        self.tenant_id = require_setting("SHAREPOINT_TENANT_ID")
        self.client_id = require_setting("SHAREPOINT_CLIENT_ID")
        self.client_secret = require_setting("SHAREPOINT_CLIENT_SECRET")
        self._token: str | None = None

    def download_from_message(self, message: dict) -> GraphFile:
        if message.get("sharePointSharingUrl"):
            return self.download_from_sharing_url(message["sharePointSharingUrl"])

        drive_path = message.get("sharePointDrivePath")
        if drive_path:
            hostname = message.get("sharePointHostname") or require_setting("SHAREPOINT_HOSTNAME")
            site_path = message.get("sharePointSitePath") or require_setting("SHAREPOINT_SITE_PATH")
            drive_id = message.get("sharePointDriveId")
            return self.download_from_drive_path(hostname, site_path, drive_path, drive_id=drive_id)

        raise ValueError(
            "Queue message must contain either sharePointSharingUrl or sharePointDrivePath."
        )

    def open_stream_from_message(self, message: dict) -> GraphFileStream:
        if message.get("sharePointSharingUrl"):
            return self.open_stream_from_sharing_url(message["sharePointSharingUrl"])

        drive_path = message.get("sharePointDrivePath")
        if drive_path:
            hostname = message.get("sharePointHostname") or require_setting("SHAREPOINT_HOSTNAME")
            site_path = message.get("sharePointSitePath") or require_setting("SHAREPOINT_SITE_PATH")
            drive_id = message.get("sharePointDriveId")
            return self.open_stream_from_drive_path(hostname, site_path, drive_path, drive_id=drive_id)

        raise ValueError(
            "Queue message must contain either sharePointSharingUrl or sharePointDrivePath."
        )

    def download_from_sharing_url(self, sharing_url: str) -> GraphFile:
        sharing_token = self._sharing_token(sharing_url)
        item = self._get(f"/shares/{sharing_token}/driveItem")
        drive_id = item["parentReference"]["driveId"]
        item_id = item["id"]
        content = self._get_bytes(f"/drives/{drive_id}/items/{item_id}/content")
        parent = item.get("parentReference", {})
        web_url = item.get("webUrl") or ""
        parsed_web_url = urlparse(web_url) if web_url else None
        return GraphFile(
            name=item.get("name", "sharepoint-file"),
            content=content,
            source_url=web_url or "graph://sharepoint-sharing-url",
            source_type="SharePoint",
            graph_client_id=self.client_id,
            mime_type=_graph_mime_type(item),
            sharepoint_hostname=parsed_web_url.netloc if parsed_web_url else None,
            sharepoint_site_id=parent.get("siteId"),
            sharepoint_drive_id=drive_id,
            sharepoint_item_id=item_id,
        )

    def open_stream_from_sharing_url(self, sharing_url: str) -> GraphFileStream:
        sharing_token = self._sharing_token(sharing_url)
        item = self._get(f"/shares/{sharing_token}/driveItem")
        drive_id = item["parentReference"]["driveId"]
        item_id = item["id"]
        response = self._get_stream(f"/drives/{drive_id}/items/{item_id}/content")
        parent = item.get("parentReference", {})
        web_url = item.get("webUrl") or ""
        parsed_web_url = urlparse(web_url) if web_url else None
        return GraphFileStream(
            name=item.get("name", "sharepoint-file"),
            content=response.raw,
            source_url=web_url or "graph://sharepoint-sharing-url",
            source_type="SharePoint",
            graph_client_id=self.client_id,
            close=response.close,
            size=item.get("size"),
            mime_type=_graph_mime_type(item),
            sharepoint_hostname=parsed_web_url.netloc if parsed_web_url else None,
            sharepoint_site_id=parent.get("siteId"),
            sharepoint_drive_id=drive_id,
            sharepoint_item_id=item_id,
        )

    def download_from_drive_path(
        self,
        hostname: str,
        site_path: str,
        drive_path: str,
        *,
        drive_id: str | None = None,
    ) -> GraphFile:
        site = None if drive_id else self._get(f"/sites/{hostname}:{site_path}")
        clean_drive_path = drive_path.strip("/")
        encoded_drive_path = quote(clean_drive_path, safe="/")

        if drive_id:
            item = self._get(f"/drives/{drive_id}/root:/{encoded_drive_path}")
            content = self._get_bytes(f"/drives/{drive_id}/items/{item['id']}/content")
        else:
            item = self._get(f"/sites/{site['id']}/drive/root:/{encoded_drive_path}")
            content = self._get_bytes(f"/sites/{site['id']}/drive/items/{item['id']}/content")

        parent = item.get("parentReference", {})
        return GraphFile(
            name=item.get("name", clean_drive_path.rsplit("/", 1)[-1]),
            content=content,
            source_url=item.get("webUrl") or f"https://{hostname}{site_path}/{clean_drive_path}",
            source_type="SharePoint",
            graph_client_id=self.client_id,
            mime_type=_graph_mime_type(item),
            sharepoint_hostname=hostname,
            sharepoint_site_path=site_path,
            sharepoint_drive_path=clean_drive_path,
            sharepoint_site_id=parent.get("siteId") or (site.get("id") if site else None),
            sharepoint_drive_id=parent.get("driveId") or drive_id,
            sharepoint_item_id=item.get("id"),
        )

    def open_stream_from_drive_path(
        self,
        hostname: str,
        site_path: str,
        drive_path: str,
        *,
        drive_id: str | None = None,
    ) -> GraphFileStream:
        site = None if drive_id else self._get(f"/sites/{hostname}:{site_path}")
        clean_drive_path = drive_path.strip("/")
        encoded_drive_path = quote(clean_drive_path, safe="/")

        if drive_id:
            item = self._get(f"/drives/{drive_id}/root:/{encoded_drive_path}")
            response = self._get_stream(f"/drives/{drive_id}/items/{item['id']}/content")
        else:
            item = self._get(f"/sites/{site['id']}/drive/root:/{encoded_drive_path}")
            response = self._get_stream(f"/sites/{site['id']}/drive/items/{item['id']}/content")

        parent = item.get("parentReference", {})
        return GraphFileStream(
            name=item.get("name", clean_drive_path.rsplit("/", 1)[-1]),
            content=response.raw,
            source_url=item.get("webUrl") or f"https://{hostname}{site_path}/{clean_drive_path}",
            source_type="SharePoint",
            graph_client_id=self.client_id,
            close=response.close,
            size=item.get("size"),
            mime_type=_graph_mime_type(item),
            sharepoint_hostname=hostname,
            sharepoint_site_path=site_path,
            sharepoint_drive_path=clean_drive_path,
            sharepoint_site_id=parent.get("siteId") or (site.get("id") if site else None),
            sharepoint_drive_id=parent.get("driveId") or drive_id,
            sharepoint_item_id=item.get("id"),
        )

    def iter_files_from_folder(
        self,
        hostname: str,
        site_path: str,
        folder_path: str,
        *,
        recursive: bool = True,
        drive_id: str | None = None,
    ) -> Iterator[dict]:
        site = None if drive_id else self._get(f"/sites/{hostname}:{site_path}")
        clean_folder_path = folder_path.strip("/")
        encoded_folder_path = quote(clean_folder_path, safe="/")

        if drive_id:
            folder = self._get(f"/drives/{drive_id}/root:/{encoded_folder_path}")
        else:
            folder = self._get(f"/sites/{site['id']}/drive/root:/{encoded_folder_path}")
            drive_id = folder.get("parentReference", {}).get("driveId")

        parent = folder.get("parentReference", {})
        site_id = parent.get("siteId") or (site.get("id") if site else None)
        yield from self._iter_children(
            site_id=site_id,
            drive_id=drive_id,
            parent_item_id=folder["id"],
            parent_drive_path=clean_folder_path,
            recursive=recursive,
            hostname=hostname,
            site_path=site_path,
        )

    def _iter_children(
        self,
        *,
        site_id: str | None,
        drive_id: str,
        parent_item_id: str,
        parent_drive_path: str,
        recursive: bool,
        hostname: str,
        site_path: str,
    ) -> Iterator[dict]:
        next_url = f"{GRAPH_ROOT}/drives/{drive_id}/items/{parent_item_id}/children"
        while next_url:
            page = self._get_url(next_url)
            for item in page.get("value", []):
                item_name = item.get("name", "")
                item_path = f"{parent_drive_path.rstrip('/')}/{item_name}".strip("/")
                parent = item.get("parentReference", {})
                item_drive_id = parent.get("driveId") or drive_id

                if "folder" in item:
                    if recursive:
                        yield from self._iter_children(
                            site_id=site_id,
                            drive_id=item_drive_id,
                            parent_item_id=item["id"],
                            parent_drive_path=item_path,
                            recursive=recursive,
                            hostname=hostname,
                            site_path=site_path,
                        )
                    continue

                if "file" not in item:
                    continue

                yield {
                    "name": item_name,
                    "size": item.get("size"),
                    "webUrl": item.get("webUrl"),
                    "contentType": _graph_mime_type(item),
                    "hasExtension": bool(PurePosixPath(item_name).suffix),
                    "sharePointHostname": hostname,
                    "sharePointSitePath": site_path,
                    "sharePointDrivePath": item_path,
                    "sharePointSiteId": site_id,
                    "sharePointDriveId": item_drive_id,
                    "sharePointItemId": item.get("id"),
                }

            next_url = page.get("@odata.nextLink")

    def _headers(self) -> dict[str, str]:
        if not self._token:
            self._token = self._get_token()
        return {"Authorization": f"Bearer {self._token}"}

    def _get_token(self) -> str:
        response = self._request(
            "POST",
            f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token",
            data={
                "client_id": self.client_id,
                "client_secret": self.client_secret,
                "scope": "https://graph.microsoft.com/.default",
                "grant_type": "client_credentials",
            },
            timeout=30,
        )
        response.raise_for_status()
        return response.json()["access_token"]

    def _get(self, path: str) -> dict:
        response = self._request("GET", f"{GRAPH_ROOT}{path}", headers=self._headers(), timeout=60)
        return response.json()

    def _get_url(self, url: str) -> dict:
        response = self._request("GET", url, headers=self._headers(), timeout=60)
        return response.json()

    def _get_bytes(self, path: str) -> bytes:
        response = self._request("GET", f"{GRAPH_ROOT}{path}", headers=self._headers(), timeout=300)
        return response.content

    def _get_stream(self, path: str) -> requests.Response:
        response = self._request(
            "GET",
            f"{GRAPH_ROOT}{path}",
            headers=self._headers(),
            stream=True,
            timeout=(30, 900),
        )
        response.raw.decode_content = True
        return response

    def _request(self, method: str, url: str, **kwargs) -> requests.Response:
        last_response: requests.Response | None = None
        for attempt in range(MAX_GRAPH_RETRIES + 1):
            response = requests.request(method, url, **kwargs)
            if response.status_code not in RETRY_STATUS_CODES:
                response.raise_for_status()
                return response

            last_response = response
            if attempt >= MAX_GRAPH_RETRIES:
                break

            retry_seconds = _retry_after_seconds(response) or DEFAULT_RETRY_SECONDS
            retry_seconds = min(retry_seconds * (attempt + 1), MAX_RETRY_SECONDS)
            response.close()
            time.sleep(retry_seconds)

        assert last_response is not None
        last_response.raise_for_status()
        return last_response

    @staticmethod
    def _sharing_token(url: str) -> str:
        raw = base64.b64encode(url.encode("utf-8")).decode("ascii")
        token = raw.rstrip("=").replace("/", "_").replace("+", "-")
        return f"u!{token}"


def _graph_mime_type(item: dict) -> str | None:
    value = (item.get("file") or {}).get("mimeType")
    return value or None


def _retry_after_seconds(response: requests.Response) -> int | None:
    retry_after = response.headers.get("Retry-After")
    if not retry_after:
        return None
    try:
        return max(1, int(retry_after))
    except ValueError:
        return None

