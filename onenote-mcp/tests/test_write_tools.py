import pytest
from mcp import Client

from onenote_mcp.config import Flags, load_settings
from onenote_mcp.server import OneNoteService, build_server
from onenote_mcp.state import EditState

from .fake_onenote import FakeOneNote
from .test_server import FakeAuth


def make(tmp_path, flags=Flags(allow_write=True)):
    fake = FakeOneNote()
    state = EditState(tmp_path / "state")
    svc = OneNoteService(
        auth_factory=lambda: FakeAuth(True), graph_factory=lambda s: fake, state_factory=lambda: state
    )
    return fake, state, build_server(svc, flags)


def text(r):
    return r.content[0].text


def test_write_scope_only_when_enabled(monkeypatch):
    assert load_settings().scopes == ["https://graph.microsoft.com/Notes.Read"]
    monkeypatch.setenv("ONENOTE_ALLOW_WRITE", "1")
    monkeypatch.setenv("ONENOTE_ALLOW_DELETE", "1")
    s = load_settings()
    assert s.scopes == ["https://graph.microsoft.com/Notes.ReadWrite"] and s.flags.allow_delete
    monkeypatch.delenv("ONENOTE_ALLOW_WRITE")
    assert not load_settings().flags.allow_delete  # delete needs write


async def test_edit_goes_to_copy_and_reports_it(tmp_path):
    fake, state, server = make(tmp_path)
    orig = fake.add_page("Plan", ["intro", "details"])
    async with Client(server) as c:
        r = await c.call_tool("replace_content", {
            "page_id": orig, "target_id": f"p:{{{orig}}}{{2}}", "content": "**New** details",
        })
        assert not r.is_error, text(r)
        copy = state.copy_of(orig)
        assert "First edit" in text(r) and copy in text(r)
        page_id, commands, _ = fake.updates[-1]
        assert page_id == copy
        assert commands == [{"target": f"p:{{{copy}}}{{2}}", "action": "replace",
                             "content": "<p><strong>New</strong> details</p>"}]

        r = await c.call_tool("add_content", {"page_id": orig, "content": "- [ ] follow up"})
        assert "First edit" not in text(r)
        assert fake.updates[-1][1][0]["action"] == "append"
        assert 'data-tag="to-do"' in fake.updates[-1][1][0]["content"]


async def test_create_page_then_edit_without_copy(tmp_path):
    fake, state, server = make(tmp_path)
    async with Client(server) as c:
        r = await c.call_tool("create_page", {"section_id": "sec1", "title": "Fresh", "content": "Hello"})
        new_id = text(r).split("id: ")[1].rstrip(").")
        r = await c.call_tool("set_page_title", {"page_id": new_id, "title": "Renamed"})
        assert not r.is_error and "First edit" not in text(r)
    assert fake.pages[new_id]["title"] == "Renamed" and len(fake.pages) == 1


async def test_insert_image_and_attachment(tmp_path):
    fake, _, server = make(tmp_path)
    img = tmp_path / "chart.png"
    img.write_bytes(b"\x89PNG")
    doc = tmp_path / "doc.pdf"
    doc.write_bytes(b"%PDF")
    orig = fake.add_page("Plan", ["a"])
    async with Client(server) as c:
        await c.call_tool("insert_image", {"page_id": orig, "source": str(img), "alt_text": "Chart",
                                           "width": 400, "position": "after", "target_id": f"p:{{{orig}}}{{1}}"})
        _, commands, parts = fake.updates[-1]
        assert commands[0]["action"] == "insert" and commands[0]["position"] == "after"
        assert 'src="name:img1"' in commands[0]["content"] and parts[0].data == b"\x89PNG"
        await c.call_tool("attach_file", {"page_id": orig, "path": str(doc)})
        assert "data-attachment" in fake.updates[-1][1][0]["content"]
        r = await c.call_tool("insert_image", {"page_id": orig, "source": str(tmp_path / "nope.png")})
        assert r.is_error and "not found" in text(r)
        r = await c.call_tool("add_content", {"page_id": orig, "content": "x", "position": "before"})
        assert r.is_error and "target_id" in text(r)


async def test_restore_and_delete_gating(tmp_path):
    fake, state, server = make(tmp_path)
    orig = fake.add_page("Plan", ["good"])
    async with Client(server) as c:
        await c.call_tool("add_content", {"page_id": orig, "content": "oops"})
        r = await c.call_tool("restore_previous_version", {"page_id": orig})
        assert "Restored" in text(r)
        tools = {t.name for t in (await c.list_tools()).tools}
    assert "delete_page" not in tools

    fake, state, server = make(tmp_path / "b", Flags(allow_write=True, allow_delete=True))
    pid = fake.add_page("Bye", [])
    async with Client(server) as c:
        r = await c.call_tool("delete_page", {"page_id": pid})
    assert fake.deleted == [pid] and "Deleted" in text(r)


async def test_copy_page_tool_and_page_markers(tmp_path):
    fake, state, server = make(tmp_path)
    orig = fake.add_page("L2", ["notes"])

    async def section_pages(section_id, max_items):
        return [{"id": pid, "title": p["title"]} for pid, p in fake.pages.items()], False

    fake.section_pages = section_pages
    async with Client(server) as c:
        tools = {t.name: t for t in (await c.list_tools()).tools}
        assert not tools["copy_page"].annotations.destructive_hint
        assert "Never changes an original page" in tools["replace_content"].description
        r = await c.call_tool("copy_page", {"page_id": orig})
        copy = state.copy_of(orig)
        assert copy in text(r) and "Open in OneNote: onenote:" in text(r)
        listing = text(await c.call_tool("list_pages", {"section_id": "sec1"}))
    assert f"original, edits go to working copy {copy}" in listing
    assert f"working copy of {orig}" in listing


async def test_copy_page_runs_in_background_and_is_retry_safe(tmp_path, monkeypatch):
    monkeypatch.setattr("onenote_mcp.server.TOOL_WAIT", 0.05)
    fake, state, server = make(tmp_path)
    fake.copy_delay = 0.3
    orig = fake.add_page("L2", ["notes"])
    async with Client(server) as c:
        first = text(await c.call_tool("copy_page", {"page_id": orig}))
        assert "still running" in first and "get_copy_status" in first
        job_id = first.split("job_id '")[1].split("'")[0]
        again = text(await c.call_tool("copy_page", {"page_id": orig}))
        assert job_id in again  # joined the running job, no second copy
        done = text(await c.call_tool("get_copy_status", {"job_id": job_id, "wait_seconds": 5}))
        assert "exact copy" in done and "Open in OneNote" in done
        third = text(await c.call_tool("copy_page", {"page_id": orig}))
        assert "already has a working copy" in third
        r = await c.call_tool("get_copy_status", {"job_id": "copy-99"})
        assert r.is_error and "without job_id" in text(r)
    assert len(fake.pages) == 2


async def test_job_list_and_unfinished_marker(tmp_path):
    from onenote_mcp.graph import GraphError

    fake, state, server = make(tmp_path)
    fake.copy_supported = False
    orig = fake.add_page("L2", ["notes"])

    async def offline(url):
        raise GraphError("connection dropped", None)

    real_get_bytes, fake.get_bytes = fake.get_bytes, offline

    async def section_pages(section_id, max_items):
        return [{"id": pid, "title": p["title"]} for pid, p in fake.pages.items()], False

    fake.section_pages = section_pages
    async with Client(server) as c:
        r = await c.call_tool("copy_page", {"page_id": orig})
        assert r.is_error and "call copy_page again" in text(r)
        jobs = text(await c.call_tool("get_copy_status", {}))
        assert "copy · failed" in jobs
        listing = text(await c.call_tool("list_pages", {"section_id": "sec1"}))
        assert "UNFINISHED COPY (failed" in listing

        fake.get_bytes = real_get_bytes  # back online: copy_page finishes the same page
        r = await c.call_tool("copy_page", {"page_id": orig})
        assert "being completed" in text(r) and "Finished copying" in text(r)
    assert len(fake.pages) == 2
