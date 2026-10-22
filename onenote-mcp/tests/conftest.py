import pytest


@pytest.fixture(autouse=True)
def isolated_env(tmp_path, monkeypatch):
    """Keep tests away from any real config or token cache."""
    monkeypatch.setenv("ONENOTE_MCP_HOME", str(tmp_path / "home"))
    monkeypatch.setenv("ONENOTE_CLIENT_ID", "00000000-0000-0000-0000-000000000000")
    for var in ("ONENOTE_AUTHORITY", "ONENOTE_TOKEN_CACHE"):
        monkeypatch.delenv(var, raising=False)
