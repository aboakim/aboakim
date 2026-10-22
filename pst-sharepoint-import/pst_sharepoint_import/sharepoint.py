"""Create SharePoint list items through Microsoft Graph."""

from __future__ import annotations

import logging
import time

import requests

from .auth import GraphAuthenticator
from .config import Settings
from .fields import build_fields, column_definition
from .pst import EmailRecord

logger = logging.getLogger("pst_sharepoint_import")

RETRYABLE_STATUS = {429, 500, 502, 503, 504}


class GraphError(Exception):
    def __init__(self, status_code: int, message: str) -> None:
        super().__init__(message)
        self.status_code = status_code


class SharePointListClient:
    def __init__(self, settings: Settings, session: requests.Session, auth: GraphAuthenticator) -> None:
        self.settings = settings
        self.session = session
        self.auth = auth

    @property
    def items_url(self) -> str:
        return (
            f"{self.settings.graph_base_url}/sites/{self.settings.site_id}"
            f"/lists/{self.settings.list_id}/items"
        )

    @property
    def columns_url(self) -> str:
        return (
            f"{self.settings.graph_base_url}/sites/{self.settings.site_id}"
            f"/lists/{self.settings.list_id}/columns"
        )

    def create_item(self, record: EmailRecord) -> dict:
        fields = build_fields(record, self.settings)
        response = self._request("POST", self.items_url, json_body={"fields": fields})
        try:
            payload = response.json()
        except Exception:
            return {}
        return payload if isinstance(payload, dict) else {}

    def iter_existing_email_ids(self, column_name: str):
        """Page every list item and yield the duplicate key column."""
        url = (
            f"{self.items_url}?$select=id"
            f"&$expand=fields($select={column_name})&$top=200"
        )
        while url:
            response = self._request(
                "GET",
                url,
                extra_headers={"Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly"},
            )
            try:
                payload = response.json()
            except Exception as exc:
                raise GraphError(response.status_code, "Graph list page was not JSON.") from exc
            if not isinstance(payload, dict):
                raise GraphError(response.status_code, "Graph list page had an unexpected shape.")
            for item in payload.get("value", []):
                fields = item.get("fields") if isinstance(item, dict) else None
                if isinstance(fields, dict) and fields.get(column_name):
                    yield str(fields[column_name])
            url = str(payload.get("@odata.nextLink") or "")

    def missing_columns(self) -> list[str]:
        """Return mapped column names that are not on the list yet."""
        existing = {name.lower() for name in self.list_column_names()}
        missing: list[str] = []
        for column_name in self.settings.field_map.values():
            if not column_name:
                continue
            if column_name.lower() not in existing:
                missing.append(column_name)
        return missing

    def list_column_names(self) -> set[str]:
        names: set[str] = set()
        url = f"{self.columns_url}?$select=name"
        while url:
            response = self._request("GET", url)
            payload = response.json()
            for column in payload.get("value", []):
                if isinstance(column, dict) and column.get("name"):
                    names.add(str(column["name"]))
            url = str(payload.get("@odata.nextLink") or "")
        return names

    def ensure_columns(self) -> None:
        """Create mapped columns that are not already on the list."""
        existing = {name.lower() for name in self.list_column_names()}
        for logical_name, column_name in self.settings.field_map.items():
            if not column_name or column_name.lower() == "title":
                continue
            if column_name.lower() in existing:
                logger.info("Column %s already exists", column_name)
                continue
            definition = column_definition(logical_name, column_name)
            if definition is None:
                continue
            try:
                self._request("POST", self.columns_url, json_body=definition)
            except GraphError as exc:
                if exc.status_code == 409:
                    logger.info("Column %s already exists", column_name)
                    existing.add(column_name.lower())
                    continue
                raise
            existing.add(column_name.lower())
            logger.info("Created column %s", column_name)

    def _request(self, method: str, url: str, *, json_body=None, extra_headers=None):
        attempts = max(1, int(self.settings.max_retries))
        attempt = 0
        token_refreshed = False
        last_network_error: Exception | None = None
        while attempt < attempts:
            attempt += 1
            headers = {
                "Authorization": f"Bearer {self.auth.get_token()}",
                "Accept": "application/json",
            }
            if json_body is not None:
                headers["Content-Type"] = "application/json"
            if extra_headers:
                headers.update(extra_headers)
            try:
                response = self.session.request(
                    method,
                    url,
                    headers=headers,
                    json=json_body,
                    timeout=self.settings.timeout_seconds,
                )
            except requests.RequestException as exc:
                last_network_error = exc
                if attempt >= attempts:
                    raise GraphError(0, f"Graph request failed: {exc}") from exc
                delay = min(60.0, float(2 ** (attempt - 1)))
                logger.warning("Graph request error (%s). Waiting %s seconds.", exc, delay)
                time.sleep(delay)
                continue

            if (
                response.status_code == 401
                and self.settings.auth_mode != "access_token"
                and not token_refreshed
            ):
                logger.warning("Graph returned 401. Refreshing the access token.")
                self.auth.invalidate()
                token_refreshed = True
                attempt -= 1
                continue

            if response.status_code in RETRYABLE_STATUS and attempt < attempts:
                delay = _retry_delay(response, attempt)
                logger.warning(
                    "Graph returned %s. Waiting %s seconds (attempt %s of %s).",
                    response.status_code,
                    delay,
                    attempt,
                    attempts,
                )
                time.sleep(delay)
                continue

            if response.status_code >= 400:
                raise GraphError(
                    response.status_code,
                    f"Graph {method} failed ({response.status_code}): {_error_text(response)}",
                )
            return response

        detail = f" Last error: {last_network_error}" if last_network_error else ""
        raise GraphError(0, f"Graph request failed after {attempts} attempts.{detail}")


def _retry_delay(response, attempt: int) -> float:
    header = ""
    headers = getattr(response, "headers", None) or {}
    header = headers.get("Retry-After", "") if hasattr(headers, "get") else ""
    if header:
        try:
            return max(0.0, float(header))
        except (TypeError, ValueError):
            pass
    return min(60.0, float(2 ** (attempt - 1)))


def _error_text(response) -> str:
    try:
        payload = response.json()
    except Exception:
        payload = None
    if isinstance(payload, dict):
        error = payload.get("error")
        if isinstance(error, dict) and error.get("message"):
            return str(error["message"])[:500]
    text = getattr(response, "text", "") or ""
    return str(text)[:500]
