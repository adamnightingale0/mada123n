"""MSAL device code sign-in with a persistent, owner-only token cache."""

from __future__ import annotations

import logging
import os
import sys
import tempfile
import threading
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

import msal

from .config import GRAPH_SCOPES, Settings

log = logging.getLogger(__name__)


class AuthError(Exception):
    """Sign-in failed or is required. The message is safe to show to the user."""


class SignInRequired(AuthError):
    """No usable token; the user has to complete a device code sign-in."""


class PersistentTokenCache(msal.SerializableTokenCache):
    """Token cache stored as a JSON file readable only by the current user.

    The file holds refresh tokens, so it is created with mode 0600 inside a
    0700 directory. On Windows, POSIX modes are mostly ignored; the default
    location under %LOCALAPPDATA% is already private to the user profile.
    """

    def __init__(self, path: Path):
        super().__init__()
        self.path = path
        # Not "_lock": msal.TokenCache already uses that name internally.
        self._file_lock = threading.Lock()
        if path.exists():
            _restrict_permissions(path)
            try:
                self.deserialize(path.read_text(encoding="utf-8"))
            except (OSError, ValueError) as exc:
                log.warning("Ignoring unreadable token cache %s: %s", path, exc)

    def save(self) -> None:
        with self._file_lock:
            if not self.has_state_changed:
                return
            _write_private(self.path, self.serialize())
            self.has_state_changed = False

    def clear(self) -> None:
        with self._file_lock:
            self.deserialize("{}")
            self.has_state_changed = False
            try:
                self.path.unlink()
            except FileNotFoundError:
                pass


def _restrict_permissions(path: Path) -> None:
    if sys.platform == "win32":
        return
    try:
        os.chmod(path, 0o600)
    except OSError as exc:
        log.warning("Could not restrict permissions on %s: %s", path, exc)


def _write_private(path: Path, data: str) -> None:
    """Atomically write ``data`` to ``path`` with owner-only permissions."""
    path.parent.mkdir(parents=True, exist_ok=True)
    if sys.platform != "win32":
        os.chmod(path.parent, 0o700)
    fd, tmp = tempfile.mkstemp(dir=path.parent, prefix=".token_cache.", suffix=".tmp")
    try:
        if sys.platform != "win32":
            os.fchmod(fd, 0o600)
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            fh.write(data)
        os.replace(tmp, path)
    except BaseException:
        try:
            os.unlink(tmp)
        except FileNotFoundError:
            pass
        raise


@dataclass
class DeviceCodePrompt:
    verification_uri: str
    user_code: str
    expires_in: int
    message: str


class Authenticator:
    """Gets Graph access tokens, running device code sign-in when needed.

    Inside the MCP server the device flow runs in a background thread: the
    first tool call returns the URL and code to the user, and later calls
    succeed once sign-in completes. The CLI ``login`` command runs it in the
    foreground instead.
    """

    def __init__(self, settings: Settings, app: Any | None = None):
        self.settings = settings
        self.cache = PersistentTokenCache(settings.token_cache_path)
        self.app = app or msal.PublicClientApplication(
            settings.client_id,
            authority=settings.authority_url,
            token_cache=self.cache,
        )
        self._lock = threading.Lock()
        self._pending: DeviceCodePrompt | None = None
        self._last_error: str | None = None

    # -- token acquisition -------------------------------------------------

    def get_token(self, force_refresh: bool = False) -> str:
        """Return an access token from the cache (refreshing silently if needed).

        Raises SignInRequired if there is no cached account or the refresh
        token no longer works.
        """
        accounts = self.app.get_accounts()
        if accounts:
            result = self.app.acquire_token_silent(
                GRAPH_SCOPES, account=accounts[0], force_refresh=force_refresh
            )
            self.cache.save()
            if result and "access_token" in result:
                return result["access_token"]
            if result and "error" in result:
                log.info("Silent token refresh failed: %s", result.get("error"))
        raise SignInRequired("Not signed in to Microsoft.")

    def signed_in_user(self) -> str | None:
        accounts = self.app.get_accounts()
        return accounts[0].get("username") if accounts else None

    # -- device code flow --------------------------------------------------

    def _initiate(self) -> tuple[dict, DeviceCodePrompt]:
        flow = self.app.initiate_device_flow(scopes=GRAPH_SCOPES)
        if "user_code" not in flow:
            raise AuthError(
                "Could not start Microsoft sign-in: "
                f"{flow.get('error_description') or flow.get('error') or flow}. "
                "Check that the client ID is right and that 'Allow public client flows' "
                "is enabled on the app registration."
            )
        prompt = DeviceCodePrompt(
            verification_uri=flow["verification_uri"],
            user_code=flow["user_code"],
            expires_in=int(flow.get("expires_in", 900)),
            message=flow["message"],
        )
        return flow, prompt

    def _complete(self, flow: dict) -> None:
        result = self.app.acquire_token_by_device_flow(flow)
        self.cache.save()
        if "access_token" not in result:
            raise AuthError(
                "Microsoft sign-in did not complete: "
                f"{result.get('error_description') or result.get('error') or 'unknown error'}"
            )

    def login_interactive(self, show: Callable[[str], None]) -> str | None:
        """Foreground device code sign-in (for the CLI). Returns the username."""
        flow, prompt = self._initiate()
        show(prompt.message)
        self._complete(flow)
        return self.signed_in_user()

    def start_background_login(self) -> DeviceCodePrompt:
        """Start (or return the already-pending) device code sign-in."""
        with self._lock:
            if self._pending is not None:
                return self._pending
            flow, prompt = self._initiate()
            self._pending = prompt
            self._last_error = None

        def worker() -> None:
            try:
                self._complete(flow)
                log.info("Microsoft sign-in completed for %s", self.signed_in_user())
            except Exception as exc:  # reported on the next tool call
                log.warning("Background sign-in failed: %s", exc)
                with self._lock:
                    self._last_error = str(exc)
            finally:
                with self._lock:
                    self._pending = None

        threading.Thread(target=worker, name="onenote-device-login", daemon=True).start()
        return prompt

    def take_last_error(self) -> str | None:
        with self._lock:
            err, self._last_error = self._last_error, None
            return err

    def logout(self) -> None:
        for account in self.app.get_accounts():
            self.app.remove_account(account)
        self.cache.clear()
