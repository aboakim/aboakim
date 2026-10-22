"""Microsoft Graph tokens for a one-time archive import.

access_token
    Use a bearer token you already minted. It will expire during a long PST run.

client_credentials
    Unattended app registration. Application permission Sites.ReadWrite.All
    (or Sites.Selected granted on this site) with admin consent.

device_code
    Interactive sign-in. Delegated permission Sites.ReadWrite.All.
    The refresh token is stored in state/token_cache.json.
"""

from __future__ import annotations

import time
from pathlib import Path

from .config import ConfigError, Settings

DELEGATED_SCOPES = ["https://graph.microsoft.com/Sites.ReadWrite.All"]
APP_SCOPES = ["https://graph.microsoft.com/.default"]


class GraphAuthenticator:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self._token = ""
        self._expires_at = 0.0

    def invalidate(self) -> None:
        self._token = ""
        self._expires_at = 0.0

    def get_token(self) -> str:
        if self._token and time.time() < self._expires_at - 60:
            return self._token
        token, expires_in = self._acquire()
        self._token = token
        self._expires_at = time.time() + max(0, expires_in)
        return token

    def _acquire(self) -> tuple[str, int]:
        mode = self.settings.auth_mode
        if mode == "access_token":
            if not self.settings.access_token:
                raise ConfigError("access_token is empty. Set GRAPH_ACCESS_TOKEN.")
            return self.settings.access_token, 300
        if mode == "client_credentials":
            return self._client_credentials()
        if mode == "device_code":
            return self._device_code()
        raise ConfigError(f"Unknown auth_mode '{mode}'.")

    def _authority(self) -> str:
        tenant = self.settings.tenant_id
        return f"https://login.microsoftonline.com/{tenant}"

    def _client_credentials(self) -> tuple[str, int]:
        msal = _load_msal()
        app = msal.ConfidentialClientApplication(
            self.settings.client_id,
            authority=self._authority(),
            client_credential=self.settings.client_secret,
        )
        result = app.acquire_token_for_client(scopes=APP_SCOPES)
        return _token_from_result(result)

    def _device_code(self) -> tuple[str, int]:
        msal = _load_msal()
        cache = msal.SerializableTokenCache()
        cache_path = Path(self.settings.token_cache_file)
        if cache_path.is_file():
            cache.deserialize(cache_path.read_text(encoding="utf-8"))
        app = msal.PublicClientApplication(
            self.settings.client_id,
            authority=self._authority(),
            token_cache=cache,
        )
        accounts = app.get_accounts()
        result = None
        if accounts:
            result = app.acquire_token_silent(DELEGATED_SCOPES, account=accounts[0])
        if not _has_access_token(result):
            flow = app.initiate_device_flow(scopes=DELEGATED_SCOPES)
            if "error" in flow:
                raise ConfigError(flow.get("error_description") or flow.get("error") or "Device code failed.")
            print(flow["message"], flush=True)
            result = app.acquire_token_by_device_flow(flow)
        token, expires_in = _token_from_result(result)
        if cache.has_state_changed:
            cache_path.parent.mkdir(parents=True, exist_ok=True)
            cache_path.write_text(cache.serialize(), encoding="utf-8")
            try:
                cache_path.chmod(0o600)
            except OSError:
                pass
        return token, expires_in


def _load_msal():
    try:
        import msal
    except ImportError as exc:
        raise ConfigError(
            "msal is not installed. Run: pip install -r requirements.txt"
        ) from exc
    return msal


def _has_access_token(result: dict | None) -> bool:
    return isinstance(result, dict) and bool(result.get("access_token"))


def _token_from_result(result: dict | None) -> tuple[str, int]:
    if not _has_access_token(result):
        detail = ""
        if isinstance(result, dict):
            detail = result.get("error_description") or result.get("error") or ""
        raise ConfigError(f"Microsoft sign-in did not return a token. {detail}".strip())
    return result["access_token"], int(result.get("expires_in") or 3600)
