import json

import pytest
from mcp import Client

from onenote_mcp.auth import DeviceCodePrompt, SignInRequired
from onenote_mcp.cli import main
from onenote_mcp.server import OneNoteService, build_server


class FakeGraph:
    async def notebooks(self):
        return [{"id": "nb1", "displayName": "Personal", "lastModifiedDateTime": "2025-03-04T10:00:00Z"}]

    async def notebook_sections(self, notebook_id):
        return [{"id": "s1", "displayName": "Ideas", "path": "Work / Ideas"}]

    async def section_pages(self, section_id, max_items):
        return [{"id": "p1", "title": "Groceries", "createdDateTime": "2025-01-01T00:00:00Z"}], True

    async def all_pages(self, max_items):
        return [
            {"id": "p1", "title": "Weekly groceries", "parentNotebook": {"displayName": "Personal"},
             "parentSection": {"displayName": "Home"}},
            {"id": "p2", "title": "Holiday plans"},
        ], False

    async def page_content(self, page_id):
        return "<html><head><title>Groceries</title></head><body><div><p data-tag='to-do'>Milk</p></div></body></html>"


class FakeAuth:
    def __init__(self, signed_in):
        self.signed_in = signed_in

    def get_token(self, force_refresh=False):
        if not self.signed_in:
            raise SignInRequired("no")
        return "t"

    def take_last_error(self):
        return None

    def start_background_login(self):
        return DeviceCodePrompt("https://microsoft.com/devicelogin", "XYZ789", 900, "msg")


def server(graph=None, signed_in=True):
    svc = OneNoteService(auth_factory=lambda: FakeAuth(signed_in), graph_factory=lambda s: graph or FakeGraph())
    return build_server(svc)


def text(result):
    return result.content[0].text


async def test_exposes_read_only_tools():
    async with Client(server()) as c:
        tools = {t.name: t for t in (await c.list_tools()).tools}
    assert set(tools) == {"list_notebooks", "list_sections", "list_pages", "get_page_content", "search_pages"}
    assert all(t.annotations.read_only_hint for t in tools.values())


async def test_tool_outputs():
    async with Client(server()) as c:
        assert "Personal (id: nb1; modified 2025-03-04)" in text(await c.call_tool("list_notebooks", {}))
        assert "Work / Ideas (id: s1" in text(await c.call_tool("list_sections", {"notebook_id": "nb1"}))
        pages = text(await c.call_tool("list_pages", {"section_id": "s1"}))
        assert "Groceries (id: p1" in pages and "more exist" in pages
        content = text(await c.call_tool("get_page_content", {"page_id": "p1"}))
        assert content.startswith("# Groceries") and "[ ] Milk" in content
        hits = text(await c.call_tool("search_pages", {"query": "GROCERIES weekly"}))
        assert "Weekly groceries — Personal / Home (id: p1" in hits and "p2" not in hits


async def test_sign_in_needed_returns_code_to_model():
    class TokenGraph(FakeGraph):
        def __init__(self, svc):
            self.svc = svc

        async def notebooks(self):
            await self.svc.token()
            return []

    svc = OneNoteService(auth_factory=lambda: FakeAuth(False), graph_factory=TokenGraph)
    async with Client(build_server(svc)) as c:
        result = await c.call_tool("list_notebooks", {})
    assert result.is_error
    assert "https://microsoft.com/devicelogin" in text(result) and "XYZ789" in text(result)


async def test_graph_errors_become_tool_errors():
    from onenote_mcp.graph import GraphError

    class Broken(FakeGraph):
        async def notebooks(self):
            raise GraphError("Microsoft Graph is throttling OneNote requests (HTTP 429).", 429)

    async with Client(server(Broken())) as c:
        result = await c.call_tool("list_notebooks", {})
    assert result.is_error and "throttling" in text(result)


async def test_missing_client_id_is_reported(monkeypatch):
    monkeypatch.delenv("ONENOTE_CLIENT_ID")
    async with Client(build_server(OneNoteService())) as c:
        result = await c.call_tool("list_notebooks", {})
    assert result.is_error and "ONENOTE_CLIENT_ID" in text(result)


def test_desktop_config_merges_existing(tmp_path, capsys):
    path = tmp_path / "claude_desktop_config.json"
    path.write_text(json.dumps({"mcpServers": {"other": {"command": "x"}}, "theme": "dark"}))
    with pytest.raises(SystemExit) as exc:
        main(["desktop-config", "--write", "--path", str(path)])
    assert exc.value.code == 0
    cfg = json.loads(path.read_text())
    assert cfg["theme"] == "dark" and "other" in cfg["mcpServers"]
    entry = cfg["mcpServers"]["onenote"]
    assert entry["args"] == ["-m", "onenote_mcp"]
    assert entry["env"]["ONENOTE_CLIENT_ID"] == "00000000-0000-0000-0000-000000000000"
    assert (tmp_path / "claude_desktop_config.json.bak").exists()
