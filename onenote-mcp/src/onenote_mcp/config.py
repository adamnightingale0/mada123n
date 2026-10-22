"""Configuration: client ID, authority and file locations.

Values come from environment variables first, then from an optional
``config.json`` in the app directory. Nothing secret is hardcoded; the client
ID of a public client app is an identifier, not a credential.
"""

from __future__ import annotations

import json
import os
import sys
from dataclasses import dataclass
from pathlib import Path

APP_NAME = "onenote-mcp"

# Personal Microsoft accounts (outlook.com, hotmail.com, live.com, ...) sign in
# through the "consumers" endpoint. It works whether the app registration is
# "personal accounts only" or "work/school + personal accounts". Use "common"
# only if you also want work/school accounts; it fails for apps registered as
# personal-only.
DEFAULT_AUTHORITY = "consumers"

GRAPH_SCOPES = ["https://graph.microsoft.com/Notes.Read"]


class ConfigError(Exception):
    """Configuration is missing or invalid."""


def app_dir() -> Path:
    """Per-user directory for the config file and token cache (outside any repo)."""
    override = os.environ.get("ONENOTE_MCP_HOME")
    if override:
        return Path(override).expanduser()
    if sys.platform == "win32":
        base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
        return Path(base) / APP_NAME
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / APP_NAME
    base = os.environ.get("XDG_CONFIG_HOME") or str(Path.home() / ".config")
    return Path(base) / APP_NAME


@dataclass(frozen=True)
class Settings:
    client_id: str
    authority: str
    token_cache_path: Path

    @property
    def authority_url(self) -> str:
        if self.authority.startswith("https://"):
            return self.authority
        return f"https://login.microsoftonline.com/{self.authority}"


def _read_config_file(path: Path) -> dict:
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ConfigError(f"Could not read {path}: {exc}") from exc
    if not isinstance(data, dict):
        raise ConfigError(f"{path} must contain a JSON object")
    return data


def load_settings() -> Settings:
    base = app_dir()
    file_cfg = _read_config_file(base / "config.json")

    client_id = (os.environ.get("ONENOTE_CLIENT_ID") or file_cfg.get("client_id") or "").strip()
    if not client_id:
        raise ConfigError(
            "No client ID configured. Set the ONENOTE_CLIENT_ID environment variable "
            f"(or add \"client_id\" to {base / 'config.json'}) to the Application (client) ID "
            "of your Azure app registration."
        )

    authority = (
        os.environ.get("ONENOTE_AUTHORITY") or file_cfg.get("authority") or DEFAULT_AUTHORITY
    ).strip()

    cache = os.environ.get("ONENOTE_TOKEN_CACHE") or file_cfg.get("token_cache")
    cache_path = Path(cache).expanduser() if cache else base / "token_cache.json"

    return Settings(client_id=client_id, authority=authority, token_cache_path=cache_path)
