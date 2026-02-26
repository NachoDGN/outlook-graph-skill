#!/usr/bin/env python3
"""Authentication manager for Outlook Graph skill.

Implements delegated auth with browser interactive login and device code fallback.
"""

from __future__ import annotations

import os
import sys
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional
from urllib.parse import urlparse

from token_store import TokenStore, TokenStoreError


class AuthConfigError(RuntimeError):
    """Raised when required auth configuration is missing or invalid."""


class DependencyError(RuntimeError):
    """Raised when runtime dependencies are missing."""


class AuthError(RuntimeError):
    """Raised when authentication fails."""


DEFAULT_SCOPES = ["User.Read", "Mail.ReadWrite", "Mail.Send"]
RESERVED_SCOPES = {"openid", "profile", "offline_access"}


@dataclass
class AuthConfig:
    client_id: str
    tenant_id: str
    redirect_uri: str
    scopes: List[str]
    profile: str
    token_store_mode: str

    @property
    def authority(self) -> str:
        return f"https://login.microsoftonline.com/{self.tenant_id}"

    @classmethod
    def from_env(cls, profile_override: Optional[str] = None) -> "AuthConfig":
        profile = profile_override or os.environ.get("OUTLOOK_PROFILE", "default")
        client_id = os.environ.get("OUTLOOK_CLIENT_ID", "").strip()
        tenant_id = os.environ.get("OUTLOOK_TENANT_ID", "common").strip() or "common"
        redirect_uri = os.environ.get(
            "OUTLOOK_REDIRECT_URI",
            "http://localhost:8765",
        ).strip()
        scopes = _parse_scopes(os.environ.get("OUTLOOK_SCOPES"))
        token_store_mode = os.environ.get("OUTLOOK_TOKEN_STORE", "auto").strip().lower() or "auto"

        if token_store_mode not in {"auto", "keyring", "file"}:
            raise AuthConfigError(
                "OUTLOOK_TOKEN_STORE must be one of: auto, keyring, file"
            )

        return cls(
            client_id=client_id,
            tenant_id=tenant_id,
            redirect_uri=redirect_uri,
            scopes=scopes,
            profile=profile,
            token_store_mode=token_store_mode,
        )


class AuthManager:
    def __init__(self, config: AuthConfig):
        self.config = config
        self._msal = None
        self._cache = None
        self._app = None

        prefer_keyring = config.token_store_mode in {"auto", "keyring"}
        require_keyring = config.token_store_mode == "keyring"
        self.store = TokenStore(
            profile=config.profile,
            prefer_keyring=prefer_keyring,
            require_keyring=require_keyring,
        )

    def status(self) -> Dict[str, Any]:
        base = {
            "profile": self.config.profile,
            "configured": bool(self.config.client_id),
            "tenant_id": self.config.tenant_id,
            "redirect_uri": self.config.redirect_uri,
            "scopes": self.config.scopes,
            "token_store_backend": self.store.backend_name(),
            "authenticated": False,
        }

        if not self.config.client_id:
            base["message"] = "OUTLOOK_CLIENT_ID is not set"
            return base

        try:
            app = self._ensure_app()
        except (DependencyError, TokenStoreError) as err:
            base["message"] = str(err)
            return base

        accounts = app.get_accounts()
        base["account_count"] = len(accounts)
        if not accounts:
            base["message"] = "No cached account for this profile"
            return base

        account = accounts[0]
        base["account"] = {
            "username": account.get("username"),
            "home_account_id": account.get("home_account_id"),
        }

        result = app.acquire_token_silent(self.config.scopes, account)
        self._persist_cache_if_changed()
        if result and "access_token" in result:
            base["authenticated"] = True
            base["expires_on"] = _epoch_to_iso8601(result.get("expires_on"))
        else:
            base["message"] = "Cached account exists but token refresh failed; run auth login"

        return base

    def login(self, method: str) -> Dict[str, Any]:
        if not self.config.client_id:
            raise AuthConfigError(
                "OUTLOOK_CLIENT_ID is required. Set it before running auth login."
            )

        app = self._ensure_app()
        normalized = (method or "browser").strip().lower()
        if normalized not in {"browser", "device"}:
            raise AuthError("Auth method must be one of: browser, device")

        if normalized == "browser":
            port = _extract_local_redirect_port(self.config.redirect_uri)
            result = app.acquire_token_interactive(
                scopes=self.config.scopes,
                prompt="select_account",
                port=port,
            )
            instructions = None
        else:
            flow = app.initiate_device_flow(scopes=self.config.scopes)
            if "user_code" not in flow:
                raise AuthError("Device code flow initialization failed")
            instructions = flow.get("message")
            if instructions:
                print(instructions, file=sys.stderr)
            result = app.acquire_token_by_device_flow(flow)

        if not result or "access_token" not in result:
            raise AuthError(_extract_auth_error(result))

        persisted_backend = self._persist_cache_if_changed()
        claims = result.get("id_token_claims", {})

        payload = {
            "profile": self.config.profile,
            "method": normalized,
            "tenant_id": claims.get("tid") or self.config.tenant_id,
            "account": {
                "username": claims.get("preferred_username") or claims.get("email") or claims.get("upn"),
                "name": claims.get("name"),
                "oid": claims.get("oid"),
            },
            "scopes": sorted(result.get("scope", "").split()),
            "expires_on": _epoch_to_iso8601(result.get("expires_on")),
            "token_store_backend": persisted_backend or self.store.backend_name(),
        }
        if instructions:
            payload["device_instructions"] = instructions
        return payload

    def logout(self) -> Dict[str, Any]:
        self.store.delete()
        self._app = None
        self._cache = None
        return {
            "profile": self.config.profile,
            "logged_out": True,
        }

    def get_access_token(self) -> str:
        if not self.config.client_id:
            raise AuthConfigError("OUTLOOK_CLIENT_ID is required")

        app = self._ensure_app()
        accounts = app.get_accounts()
        if not accounts:
            raise AuthError(
                f"No cached account for profile '{self.config.profile}'. Run auth login first."
            )

        result = app.acquire_token_silent(self.config.scopes, accounts[0])
        self._persist_cache_if_changed()

        if result and "access_token" in result:
            return result["access_token"]

        raise AuthError(
            "Could not acquire access token silently. Run auth login to refresh consent."
        )

    def _ensure_app(self):
        if self._app is not None:
            return self._app

        try:
            import msal  # type: ignore
        except Exception as err:
            raise DependencyError(
                "Missing dependency 'msal'. Install with: python3 -m pip install --user msal"
            ) from err

        self._msal = msal
        self._cache = msal.SerializableTokenCache()
        serialized = self.store.load()
        if serialized:
            try:
                self._cache.deserialize(serialized)
            except Exception:
                # Corrupt cache should not hard-fail auth.
                self._cache = msal.SerializableTokenCache()

        self._app = msal.PublicClientApplication(
            client_id=self.config.client_id,
            authority=self.config.authority,
            token_cache=self._cache,
        )
        return self._app

    def _persist_cache_if_changed(self) -> Optional[str]:
        if self._cache is None:
            return None
        if not self._cache.has_state_changed:
            return None
        return self.store.save(self._cache.serialize())


def _parse_scopes(raw: Optional[str]) -> List[str]:
    if raw is None:
        return list(DEFAULT_SCOPES)

    parts: List[str] = []
    seen = set()
    for chunk in raw.replace(",", " ").split():
        scope = chunk.strip()
        if not scope:
            continue
        if scope.lower() in RESERVED_SCOPES:
            continue
        lowered = scope.lower()
        if lowered in seen:
            continue
        seen.add(lowered)
        parts.append(scope)

    return parts or list(DEFAULT_SCOPES)


def _extract_local_redirect_port(redirect_uri: str) -> int:
    parsed = urlparse(redirect_uri)
    if parsed.scheme not in {"http", "https"}:
        raise AuthConfigError(
            "OUTLOOK_REDIRECT_URI must be an http(s) URL for browser auth"
        )
    if parsed.hostname not in {"localhost", "127.0.0.1"}:
        raise AuthConfigError(
            "OUTLOOK_REDIRECT_URI host must be localhost or 127.0.0.1 for browser auth"
        )

    if parsed.port is not None:
        return parsed.port
    if parsed.scheme == "https":
        return 443
    return 80


def _epoch_to_iso8601(value: Any) -> Optional[str]:
    if value is None:
        return None
    try:
        epoch = int(value)
    except Exception:
        return None
    dt = datetime.fromtimestamp(epoch, tz=timezone.utc)
    return dt.isoformat()


def _extract_auth_error(result: Optional[Dict[str, Any]]) -> str:
    if not result:
        return "Authentication failed with an empty response"
    if "error_description" in result:
        return str(result["error_description"])
    if "error" in result:
        return str(result["error"])
    return "Authentication failed"
