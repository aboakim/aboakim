"""Load importer settings from a JSON file and the environment."""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass, field
from pathlib import Path

AUTH_MODES = ("access_token", "client_credentials", "device_code")
COLUMN_NAME = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")

ENV_OVERRIDES = {
    "PST_FILE": "pst_file",
    "SHAREPOINT_SITE_ID": "site_id",
    "SHAREPOINT_LIST_ID": "list_id",
    "GRAPH_AUTH_MODE": "auth_mode",
    "GRAPH_TENANT_ID": "tenant_id",
    "GRAPH_CLIENT_ID": "client_id",
    "GRAPH_CLIENT_SECRET": "client_secret",
    "GRAPH_ACCESS_TOKEN": "access_token",
}


class ConfigError(Exception):
    """The import cannot start because configuration is incomplete."""


from .fields import DEFAULT_FIELD_MAP  # noqa: E402  # after ConfigError so PST import can use it


@dataclass
class Limits:
    subject: int = 255
    email_id: int = 64
    internet_message_id: int = 255
    pst_identifier: int = 64
    sender: int = 255
    sender_name: int = 255
    recipients: int = 4000
    folder: int = 1000
    body: int = 60000
    body_source: int = 32
    attachments: int = 4000
    source_file: int = 255


@dataclass
class Settings:
    pst_file: str
    site_id: str = ""
    list_id: str = ""
    graph_base_url: str = "https://graph.microsoft.com/v1.0"
    auth_mode: str = "device_code"
    tenant_id: str = ""
    client_id: str = ""
    client_secret: str = ""
    access_token: str = ""
    field_map: dict[str, str] = field(default_factory=lambda: dict(DEFAULT_FIELD_MAP))
    limits: Limits = field(default_factory=Limits)
    request_delay_seconds: float = 0.2
    max_retries: int = 5
    timeout_seconds: int = 60
    dry_run: bool = False
    limit: int | None = None
    sync_existing: bool = True
    ensure_columns: bool = False
    folder_prefixes: list[str] = field(default_factory=list)
    state_file: str = "state/imported.jsonl"
    log_dir: str = "logs"
    log_level: str = "INFO"
    log_each_message: bool = False
    token_cache_file: str = "state/token_cache.json"
    config_dir: str = ""

    def needs_graph(self) -> bool:
        return self.ensure_columns or self.sync_existing or not self.dry_run

    def validate(self) -> None:
        if not self.pst_file:
            raise ConfigError("pst_file is required.")
        pst_path = Path(self.pst_file)
        if not pst_path.is_file():
            raise ConfigError(f"PST file not found: {self.pst_file}")
        if self.auth_mode not in AUTH_MODES:
            allowed = ", ".join(AUTH_MODES)
            raise ConfigError(f"auth_mode must be one of: {allowed}.")
        if self.limit is not None and self.limit < 1:
            raise ConfigError("limit must be a positive integer.")
        if self.max_retries < 1:
            raise ConfigError("max_retries must be at least 1.")
        self._validate_field_map()
        if self.needs_graph():
            if not self.site_id or not self.list_id:
                raise ConfigError(
                    "site_id and list_id are required. "
                    "A PST-only preview can use --dry-run --no-sync-existing."
                )
            self._validate_auth()

    def _validate_field_map(self) -> None:
        for logical_name, column_name in self.field_map.items():
            if logical_name not in DEFAULT_FIELD_MAP:
                raise ConfigError(f"Unknown field_map key: {logical_name}")
            if column_name and not COLUMN_NAME.match(column_name):
                raise ConfigError(
                    f"SharePoint column '{column_name}' for {logical_name} "
                    "must be an internal name of letters, numbers, and underscores."
                )
        if not self.field_map.get("email_id"):
            raise ConfigError("field_map.email_id is required so duplicates can be skipped.")
        if not self.field_map.get("subject"):
            raise ConfigError("field_map.subject is required. The default column is Title.")

    def _validate_auth(self) -> None:
        if self.auth_mode == "access_token" and not self.access_token:
            raise ConfigError("access_token is empty. Set GRAPH_ACCESS_TOKEN.")
        if self.auth_mode == "client_credentials":
            missing = [
                name
                for name in ("tenant_id", "client_id", "client_secret")
                if not getattr(self, name)
            ]
            if missing:
                raise ConfigError(
                    "client_credentials auth needs: " + ", ".join(missing) + "."
                )
        if self.auth_mode == "device_code":
            missing = [name for name in ("tenant_id", "client_id") if not getattr(self, name)]
            if missing:
                raise ConfigError("device_code auth needs: " + ", ".join(missing) + ".")


def load_env_file(path: Path) -> None:
    """Load KEY=VALUE lines. Existing environment variables win."""
    if not path.is_file():
        return
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key:
            os.environ.setdefault(key, value)


def load_settings(path: str | Path) -> Settings:
    config_path = Path(path).resolve()
    if not config_path.is_file():
        raise ConfigError(f"Config file not found: {config_path}")
    load_env_file(config_path.parent / ".env")

    try:
        data = json.loads(config_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ConfigError(f"Config file is not valid JSON: {exc}") from exc
    if not isinstance(data, dict):
        raise ConfigError("Config file must be a JSON object.")

    settings = Settings(pst_file=str(data.get("pst_file", "") or ""))
    settings.config_dir = str(config_path.parent)
    _apply_json(settings, data)
    _apply_environment(settings)
    _resolve_local_paths(settings, config_path.parent)
    return settings


def _apply_json(settings: Settings, data: dict) -> None:
    string_fields = (
        "pst_file",
        "site_id",
        "list_id",
        "graph_base_url",
        "auth_mode",
        "tenant_id",
        "client_id",
        "client_secret",
        "access_token",
        "state_file",
        "log_dir",
        "log_level",
        "token_cache_file",
    )
    for name in string_fields:
        if name in data and data[name] is not None:
            setattr(settings, name, str(data[name]))
    if "field_map" in data:
        if not isinstance(data["field_map"], dict):
            raise ConfigError("field_map must be an object.")
        merged = dict(DEFAULT_FIELD_MAP)
        for key, value in data["field_map"].items():
            merged[str(key)] = "" if value is None else str(value)
        settings.field_map = merged
    if "limits" in data:
        if not isinstance(data["limits"], dict):
            raise ConfigError("limits must be an object.")
        for key, value in data["limits"].items():
            if not hasattr(settings.limits, key):
                raise ConfigError(f"Unknown limit: {key}")
            setattr(settings.limits, key, int(value))
    if "folder_prefixes" in data:
        prefixes = data["folder_prefixes"]
        if not isinstance(prefixes, list):
            raise ConfigError("folder_prefixes must be a list of strings.")
        settings.folder_prefixes = [str(item) for item in prefixes]
    for name in (
        "request_delay_seconds",
        "max_retries",
        "timeout_seconds",
        "dry_run",
        "sync_existing",
        "ensure_columns",
        "log_each_message",
    ):
        if name in data and data[name] is not None:
            setattr(settings, name, data[name])
    if "limit" in data and data["limit"] is not None:
        settings.limit = int(data["limit"])
    settings.request_delay_seconds = float(settings.request_delay_seconds)
    settings.max_retries = int(settings.max_retries)
    settings.timeout_seconds = int(settings.timeout_seconds)
    settings.dry_run = bool(settings.dry_run)
    settings.sync_existing = bool(settings.sync_existing)
    settings.ensure_columns = bool(settings.ensure_columns)
    settings.log_each_message = bool(settings.log_each_message)


def _apply_environment(settings: Settings) -> None:
    for env_name, attr in ENV_OVERRIDES.items():
        value = os.environ.get(env_name)
        if value:
            setattr(settings, attr, value)


def _resolve_local_paths(settings: Settings, config_dir: Path) -> None:
    settings.state_file = str(_resolve(config_dir, settings.state_file))
    settings.log_dir = str(_resolve(config_dir, settings.log_dir))
    settings.token_cache_file = str(_resolve(config_dir, settings.token_cache_file))


def _resolve(base: Path, value: str) -> Path:
    path = Path(value)
    if path.is_absolute():
        return path
    return (base / path).resolve()
