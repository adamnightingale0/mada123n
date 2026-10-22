import json
import os
import stat
import sys
import threading

import pytest

from onenote_mcp.auth import AuthError, Authenticator, PersistentTokenCache, SignInRequired
from onenote_mcp.config import ConfigError, load_settings


class FakeApp:
    """Stands in for msal.PublicClientApplication."""

    def __init__(self, cache, accounts=None, silent=None, device_result=None, flow=None):
        self.cache = cache
        self.accounts = accounts or []
        self.silent = silent
        self.device_result = device_result or {"access_token": "new"}
        self.flow = flow or {
            "user_code": "ABC123",
            "verification_uri": "https://microsoft.com/devicelogin",
            "expires_in": 900,
            "message": "Go to https://microsoft.com/devicelogin and enter ABC123",
        }
        self.release = threading.Event()
        self.release.set()

    def get_accounts(self):
        return self.accounts

    def acquire_token_silent(self, scopes, account, force_refresh=False):
        return self.silent

    def initiate_device_flow(self, scopes):
        assert scopes == ["https://graph.microsoft.com/Notes.Read"]
        return self.flow

    def acquire_token_by_device_flow(self, flow):
        self.release.wait(5)
        if "access_token" in self.device_result:
            self.accounts = [{"username": "me@outlook.com"}]
            self.cache.has_state_changed = True
        return self.device_result

    def remove_account(self, account):
        self.accounts.remove(account)


def make_auth(**kw):
    settings = load_settings()
    auth = Authenticator(settings, app=object())
    auth.app = FakeApp(auth.cache, **kw)
    return auth


def test_settings_default_to_consumers_authority():
    s = load_settings()
    assert s.authority_url == "https://login.microsoftonline.com/consumers"


def test_settings_require_client_id(monkeypatch):
    monkeypatch.delenv("ONENOTE_CLIENT_ID")
    with pytest.raises(ConfigError, match="ONENOTE_CLIENT_ID"):
        load_settings()


def test_settings_read_config_file(monkeypatch, tmp_path):
    monkeypatch.delenv("ONENOTE_CLIENT_ID")
    home = tmp_path / "home"
    home.mkdir()
    (home / "config.json").write_text(json.dumps({"client_id": "abc", "authority": "common"}))
    s = load_settings()
    assert (s.client_id, s.authority_url) == ("abc", "https://login.microsoftonline.com/common")


@pytest.mark.skipif(sys.platform == "win32", reason="POSIX permissions")
def test_token_cache_is_owner_only(tmp_path):
    path = tmp_path / "d" / "token_cache.json"
    cache = PersistentTokenCache(path)
    cache.has_state_changed = True
    cache.save()
    assert stat.S_IMODE(os.stat(path).st_mode) == 0o600
    assert stat.S_IMODE(os.stat(path.parent).st_mode) == 0o700
    os.chmod(path, 0o644)
    PersistentTokenCache(path)  # loading re-tightens
    assert stat.S_IMODE(os.stat(path).st_mode) == 0o600


def test_silent_token():
    auth = make_auth(accounts=[{"username": "me@outlook.com"}], silent={"access_token": "t"})
    assert auth.get_token() == "t"


def test_no_account_requires_sign_in():
    with pytest.raises(SignInRequired):
        make_auth().get_token()


def test_failed_refresh_requires_sign_in():
    auth = make_auth(accounts=[{"username": "x"}], silent={"error": "invalid_grant"})
    with pytest.raises(SignInRequired):
        auth.get_token()


def test_background_login_returns_same_prompt_until_done():
    auth = make_auth()
    auth.app.release.clear()
    first = auth.start_background_login()
    second = auth.start_background_login()
    assert first is second and first.user_code == "ABC123"
    auth.app.release.set()
    for t in threading.enumerate():
        if t.name == "onenote-device-login":
            t.join(5)
    assert auth.signed_in_user() == "me@outlook.com"
    assert auth.settings.token_cache_path.exists()


def test_background_login_failure_is_reported():
    auth = make_auth(device_result={"error": "authorization_declined", "error_description": "declined"})
    auth.start_background_login()
    for t in threading.enumerate():
        if t.name == "onenote-device-login":
            t.join(5)
    assert "declined" in auth.take_last_error()
    assert auth.take_last_error() is None


def test_bad_client_id_gives_clear_error():
    auth = make_auth(flow={"error": "unauthorized_client", "error_description": "bad app"})
    with pytest.raises(AuthError, match="public client flows"):
        auth.start_background_login()
