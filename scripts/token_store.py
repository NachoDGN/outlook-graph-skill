#!/usr/bin/env python3
"""Token persistence for outlook-graph.

Keyring is preferred when available. A strict-permission file fallback is used
when keyring is unavailable or explicitly disabled.
"""

from __future__ import annotations

import json
import os
import stat
from pathlib import Path
from typing import Optional


class TokenStoreError(RuntimeError):
    """Raised when token cache persistence fails."""


class TokenStore:
    def __init__(
        self,
        profile: str,
        service_name: str = "codex-outlook-graph",
        base_dir: Optional[Path] = None,
        prefer_keyring: bool = True,
        require_keyring: bool = False,
    ) -> None:
        self.profile = self._sanitize_profile(profile)
        self.service_name = service_name
        self.base_dir = base_dir or self._default_base_dir()
        self.prefer_keyring = prefer_keyring
        self.require_keyring = require_keyring
        self._keyring = self._load_keyring() if prefer_keyring else None

        if self.require_keyring and self._keyring is None:
            raise TokenStoreError(
                "OUTLOOK_TOKEN_STORE=keyring requested but no keyring backend is available. "
                "Install 'keyring' and configure an OS keychain backend, or set OUTLOOK_TOKEN_STORE=file."
            )

    @staticmethod
    def _default_base_dir() -> Path:
        override = os.environ.get("OUTLOOK_TOKEN_CACHE_DIR")
        if override:
            return Path(override).expanduser()
        return Path.home() / ".codex" / "outlook-graph" / "tokens"

    @staticmethod
    def _sanitize_profile(profile: str) -> str:
        profile = (profile or "default").strip()
        if not profile:
            return "default"
        safe = []
        for char in profile:
            if char.isalnum() or char in {"-", "_", "."}:
                safe.append(char)
            else:
                safe.append("_")
        sanitized = "".join(safe).strip("._")
        return sanitized or "default"

    def backend_name(self) -> str:
        if self._keyring is not None:
            return "keyring"
        return "file"

    def load(self) -> Optional[str]:
        if self._keyring is not None:
            try:
                value = self._keyring.get_password(self.service_name, self.profile)
                if value:
                    return value
            except Exception:
                # Fall back to file cache.
                pass
        return self._load_from_file()

    def save(self, serialized_cache: str) -> str:
        if self._keyring is not None:
            try:
                self._keyring.set_password(self.service_name, self.profile, serialized_cache)
                return "keyring"
            except Exception:
                if self.require_keyring:
                    raise TokenStoreError("Failed to write token cache to keyring backend")
                # Fall back to file cache.

        self._save_to_file(serialized_cache)
        return "file"

    def delete(self) -> None:
        if self._keyring is not None:
            try:
                self._keyring.delete_password(self.service_name, self.profile)
            except Exception:
                pass
        path = self._cache_file_path()
        if path.exists():
            path.unlink()

    def _cache_file_path(self) -> Path:
        return self.base_dir / f"{self.profile}.json"

    def _load_from_file(self) -> Optional[str]:
        path = self._cache_file_path()
        if not path.exists():
            return None

        raw = path.read_text(encoding="utf-8")
        if not raw.strip():
            return None

        try:
            payload = json.loads(raw)
        except json.JSONDecodeError:
            # Legacy/plain payload support.
            return raw

        cache = payload.get("cache")
        if isinstance(cache, str) and cache:
            return cache
        return None

    def _save_to_file(self, serialized_cache: str) -> None:
        self.base_dir.mkdir(parents=True, exist_ok=True)
        try:
            self.base_dir.chmod(0o700)
        except OSError:
            pass

        path = self._cache_file_path()
        payload = {
            "cache": serialized_cache,
            "profile": self.profile,
        }

        tmp_path = path.with_suffix(".tmp")
        tmp_path.write_text(json.dumps(payload, separators=(",", ":")), encoding="utf-8")
        try:
            tmp_path.chmod(0o600)
        except OSError:
            pass
        tmp_path.replace(path)

        # Best effort hardening in case the file already existed with loose perms.
        try:
            mode = stat.S_IMODE(path.stat().st_mode)
            if mode & 0o077:
                path.chmod(0o600)
        except OSError:
            pass

    @staticmethod
    def _load_keyring():
        try:
            import keyring  # type: ignore
        except Exception:
            return None

        try:
            backend = keyring.get_keyring()
        except Exception:
            return None

        if backend is None:
            return None

        backend_name = backend.__class__.__name__.lower()
        module_name = backend.__class__.__module__.lower()
        if "fail" in backend_name or "fail" in module_name:
            return None

        return keyring
