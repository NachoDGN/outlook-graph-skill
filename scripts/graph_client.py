#!/usr/bin/env python3
"""Microsoft Graph API client for Outlook mailbox operations."""

from __future__ import annotations

import time
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import quote, urljoin

from auth_manager import AuthManager, DependencyError


class GraphAPIError(RuntimeError):
    """Raised on Graph API failures."""

    def __init__(self, message: str, status_code: Optional[int] = None):
        super().__init__(message)
        self.status_code = status_code


class GraphClient:
    def __init__(
        self,
        auth: AuthManager,
        base_url: str = "https://graph.microsoft.com/v1.0",
        timeout_seconds: int = 30,
        max_retries: int = 2,
    ) -> None:
        self.auth = auth
        self.base_url = base_url.rstrip("/")
        self.timeout_seconds = timeout_seconds
        self.max_retries = max_retries
        self._requests = self._load_requests()

    def list_messages(
        self,
        folder: str,
        unread_only: bool,
        top: int,
        select_fields: Optional[List[str]] = None,
        has_attachments_only: bool = False,
        received_since: Optional[str] = None,
        max_pages: Optional[int] = None,
    ) -> List[Dict[str, Any]]:
        limit = max(1, min(int(top), 5000))
        page_size = min(limit, 50)

        params: Dict[str, Any] = {
            "$top": str(page_size),
            "$orderby": "receivedDateTime DESC",
        }
        filters: List[str] = []
        if unread_only:
            filters.append("isRead eq false")
        if has_attachments_only:
            filters.append("hasAttachments eq true")
        if received_since:
            filters.append(f"receivedDateTime ge {received_since}")
        if filters:
            params["$filter"] = " and ".join(filters)
        if select_fields:
            params["$select"] = ",".join(select_fields)

        path = f"/me/mailFolders/{_quote_segment(folder)}/messages"
        return self._collect_paginated(path, params=params, limit=limit, max_pages=max_pages)

    def get_mail_folder(self, folder_token: str) -> Dict[str, Any]:
        return self._request_json(
            "GET",
            f"/me/mailFolders/{_quote_segment(folder_token)}",
            params={"$select": _mail_folder_select_fields()},
        )

    def list_child_folders(self, folder_token: str, include_hidden: bool = False) -> List[Dict[str, Any]]:
        params: Dict[str, Any] = {
            "$select": _mail_folder_select_fields(),
            "$top": "50",
        }
        if include_hidden:
            params["includeHiddenFolders"] = "true"
        path = f"/me/mailFolders/{_quote_segment(folder_token)}/childFolders"
        return self._collect_paginated(path, params=params, limit=5000, max_pages=200)

    def get_message(self, message_id: str) -> Dict[str, Any]:
        return self._request_json("GET", f"/me/messages/{_quote_segment(message_id)}")

    def mark_message(self, message_id: str, read: bool) -> Dict[str, Any]:
        return self._request_json(
            "PATCH",
            f"/me/messages/{_quote_segment(message_id)}",
            json_body={"isRead": bool(read)},
        )

    def create_draft(
        self,
        to_recipients: List[str],
        subject: str,
        body: str,
        body_content_type: str = "Text",
    ) -> Dict[str, Any]:
        recipients = [
            {"emailAddress": {"address": address}}
            for address in to_recipients
        ]

        payload = {
            "subject": subject,
            "body": {
                "contentType": body_content_type,
                "content": body,
            },
            "toRecipients": recipients,
        }
        return self._request_json("POST", "/me/messages", json_body=payload)

    def send_draft(self, message_id: str) -> None:
        self._request_json("POST", f"/me/messages/{_quote_segment(message_id)}/send")

    def list_attachments(self, message_id: str) -> List[Dict[str, Any]]:
        data = self._request_json("GET", f"/me/messages/{_quote_segment(message_id)}/attachments")
        return data.get("value", [])

    def get_attachment(self, message_id: str, attachment_id: str) -> Dict[str, Any]:
        return self._request_json(
            "GET",
            f"/me/messages/{_quote_segment(message_id)}/attachments/{_quote_segment(attachment_id)}",
        )

    def download_attachment_bytes(
        self,
        message_id: str,
        attachment_id: str,
    ) -> Tuple[bytes, str]:
        return self._request_bytes(
            "GET",
            f"/me/messages/{_quote_segment(message_id)}/attachments/{_quote_segment(attachment_id)}/$value",
        )

    def _collect_paginated(
        self,
        path: str,
        params: Optional[Dict[str, Any]],
        limit: int,
        max_pages: Optional[int] = None,
    ) -> List[Dict[str, Any]]:
        items: List[Dict[str, Any]] = []
        next_url: Optional[str] = path
        next_params = dict(params or {})
        page_count = 0

        while next_url and len(items) < limit:
            if max_pages is not None and page_count >= max_pages:
                break
            payload = self._request_json("GET", next_url, params=next_params, absolute_url=next_url.startswith("http"))
            page_count += 1
            values = payload.get("value", [])
            for value in values:
                if len(items) >= limit:
                    break
                items.append(value)

            next_url = payload.get("@odata.nextLink")
            next_params = {}

        return items

    def _request_json(
        self,
        method: str,
        path_or_url: str,
        params: Optional[Dict[str, Any]] = None,
        json_body: Optional[Dict[str, Any]] = None,
        absolute_url: bool = False,
    ) -> Dict[str, Any]:
        response = self._request(method, path_or_url, params=params, json_body=json_body, absolute_url=absolute_url)
        if response.status_code == 204 or not response.content:
            return {}

        content_type = response.headers.get("Content-Type", "")
        if "application/json" not in content_type.lower():
            raise GraphAPIError(
                f"Expected JSON response but got content type '{content_type}'",
                status_code=response.status_code,
            )

        return response.json()

    def _request_bytes(
        self,
        method: str,
        path_or_url: str,
        params: Optional[Dict[str, Any]] = None,
        absolute_url: bool = False,
    ) -> Tuple[bytes, str]:
        response = self._request(method, path_or_url, params=params, absolute_url=absolute_url)
        content_type = response.headers.get("Content-Type", "application/octet-stream")
        return response.content, content_type

    def _request(
        self,
        method: str,
        path_or_url: str,
        params: Optional[Dict[str, Any]] = None,
        json_body: Optional[Dict[str, Any]] = None,
        absolute_url: bool = False,
    ):
        token = self.auth.get_access_token()
        url = path_or_url if absolute_url else self._build_url(path_or_url)
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
            "Prefer": 'IdType="ImmutableId"',
        }
        if json_body is not None:
            headers["Content-Type"] = "application/json"

        response = None
        for attempt in range(self.max_retries + 1):
            response = self._requests.request(
                method=method,
                url=url,
                params=params,
                json=json_body,
                headers=headers,
                timeout=self.timeout_seconds,
            )
            if response.status_code in {429, 500, 502, 503, 504} and attempt < self.max_retries:
                delay = _retry_delay_seconds(response, attempt)
                time.sleep(delay)
                continue
            break

        if response is None:
            raise GraphAPIError("No response from Graph API")

        if response.status_code >= 400:
            raise GraphAPIError(_extract_graph_error(response), status_code=response.status_code)

        return response

    def _build_url(self, path: str) -> str:
        normalized = path if path.startswith("/") else f"/{path}"
        return urljoin(self.base_url + "/", normalized.lstrip("/"))

    @staticmethod
    def _load_requests():
        try:
            import requests  # type: ignore
        except Exception as err:
            raise DependencyError(
                "Missing dependency 'requests'. Install with: python3 -m pip install requests (or use --user only outside virtualenv)"
            ) from err

        return requests


def _retry_delay_seconds(response: Any, attempt: int) -> float:
    retry_after = response.headers.get("Retry-After")
    if retry_after:
        try:
            return max(0.5, float(retry_after))
        except ValueError:
            pass
    # 1.0, 2.0, 4.0...
    return float(2 ** attempt)


def _extract_graph_error(response: Any) -> str:
    prefix = f"Graph API request failed with status {response.status_code}"
    try:
        payload = response.json()
    except Exception:
        body = (response.text or "").strip()
        if body:
            return f"{prefix}: {body[:500]}"
        return prefix

    error = payload.get("error") if isinstance(payload, dict) else None
    if isinstance(error, dict):
        code = error.get("code")
        message = error.get("message")
        if code and message:
            return f"{prefix}: {code} - {message}"
        if message:
            return f"{prefix}: {message}"

    return prefix


def _quote_segment(value: str) -> str:
    return quote(str(value), safe="")


def _mail_folder_select_fields() -> str:
    return "id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount,isHidden"
