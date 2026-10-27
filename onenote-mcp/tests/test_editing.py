import asyncio

import pytest

from onenote_mcp.editing import COPY_SUFFIX, EditError, Editor, find_placeholders
from onenote_mcp.state import EditState

from .fake_onenote import GRAPH, FakeOneNote


@pytest.fixture
def setup(tmp_path):
    fake = FakeOneNote()
    state = EditState(tmp_path)
    return fake, state, Editor(fake, state)


async def finish(ed, job):
    await ed.wait(job, 5)
    assert job.status == "done", job.summary()
    return job


async def test_first_edit_copies_and_later_edits_reuse(setup, tmp_path):
    fake, state, ed = setup
    orig = fake.add_page("Notes", ["one", "two"])
    prep = await ed.prepare(orig, None)
    copy = prep.page_id
    assert copy != orig and fake.pages[copy]["title"] == "Notes" + COPY_SUFFIX
    assert fake.pages[orig]["title"] == "Notes"  # original untouched
    assert "First edit" in prep.notes[0] and "exact copy" in prep.notes[0]

    again = await ed.prepare(orig, None)  # via the original's ID
    assert again.page_id == copy and "existing working copy" in again.notes[0]
    direct = await ed.prepare(copy, None)  # via the copy's ID
    assert direct.page_id == copy and direct.notes == []
    assert len(fake.pages) == 2
    assert EditState(tmp_path).copy_of(orig) == copy  # survives a restart


async def test_snapshot_is_overwritten_each_edit(setup):
    fake, state, ed = setup
    orig = fake.add_page("Notes", ["first-version"])
    copy = (await ed.prepare(orig, None)).page_id
    assert "first-version" in state.load_snapshot(copy)
    fake.pages[copy]["paragraphs"] = ["second-version"]
    await ed.prepare(copy, None)
    snap = state.load_snapshot(copy)
    assert "second-version" in snap and "first-version" not in snap
    assert len(list(state.snapshot_dir.iterdir())) == 1


async def test_target_from_original_maps_onto_copy(setup):
    fake, _, ed = setup
    orig = fake.add_page("Notes", ["alpha", "beta"])
    prep = await ed.prepare(orig, f"p:{{{orig}}}{{2}}")
    assert prep.target == f"p:{{{prep.page_id}}}{{2}}"


async def test_unknown_target_is_reported(setup):
    fake, _, ed = setup
    orig = fake.add_page("Notes", ["alpha"])
    with pytest.raises(EditError, match="get_page_structure"):
        await ed.prepare(orig, "p:{nope}{9}")


async def test_deleted_copy_is_recreated(setup):
    fake, _, ed = setup
    orig = fake.add_page("Notes", ["a"])
    first = (await ed.prepare(orig, None)).page_id
    del fake.pages[first]
    prep = await ed.prepare(orig, None)
    assert prep.page_id != first and "previous working copy was deleted" in prep.notes[0]


async def test_rebuild_creates_page_first_then_fills_placeholders(tmp_path):
    fake = FakeOneNote(copy_supported=False)
    state = EditState(tmp_path)
    ed = Editor(fake, state)
    orig = fake.add_page("L2", ["notes"])
    job = await finish(ed, ed.start_copy(orig))
    assert job.method == "rebuilt" and "ink" in job.message and "2 of 2" in job.message
    section, html, parts = fake.created[0]
    # Created without any uploads: just placeholders where the PDF and slide were.
    assert parts == [] and section == "sec1" and ' id="' not in html and 'name="created"' not in html
    assert "[Copying file 1 of 2…]" in html and "[Copying image 2 of 2…]" in html
    # Placeholders are replaced via OneNote's generated ids (data-id is refused for <p>).
    copy = job.page_id
    fills = [(cmds, up) for pid, cmds, up in fake.updates if pid == copy and cmds[0]["target"] != "title"]
    cmds, uploaded = fills[0]
    assert [c["target"] for c in cmds] == [f"p:{{{copy}}}{{2}}", f"p:{{{copy}}}{{3}}"]
    assert 'data-attachment="Topic_2.pdf"' in cmds[0]["content"] and 'data="name:file1"' in cmds[0]["content"]
    assert 'src="name:img2"' in cmds[1]["content"] and 'alt="pic"' in cmds[1]["content"]
    assert {p.name: p.data for p in uploaded} == {"file1": b"%PDF", "img2": b"\x89PNG"}
    assert fake.pages[copy]["title"] == "L2 (copy)"  # title metadata set explicitly
    assert state.copy_of(orig) == copy
    assert state.notes.get("exact_copy_unsupported") is None  # 400, not 501


async def test_501_is_remembered_and_exact_copy_skipped(tmp_path):
    from onenote_mcp.graph import GraphError

    fake = FakeOneNote()
    calls = []

    async def unsupported(page_id, section_id, timeout=90.0, on_wait=None):
        calls.append(page_id)
        raise GraphError("Not supported (HTTP 501). Details: OData Feature not implemented.", 501)

    fake.copy_page = unsupported
    state = EditState(tmp_path)
    ed = Editor(fake, state)
    job = await finish(ed, ed.start_copy(fake.add_page("A", ["x"])))
    assert "isn't available for this account" in job.message and state.notes["exact_copy_unsupported"]
    await finish(ed, Editor(fake, EditState(tmp_path)).start_copy(fake.add_page("B", ["y"])))
    assert len(calls) == 1  # remembered across restarts


async def test_rebuild_marks_images_that_fail(tmp_path):
    fake = FakeOneNote(copy_supported=False, fail_downloads={"r1"})
    ed = Editor(fake, EditState(tmp_path))
    job = await finish(ed, ed.start_copy(fake.add_page("L2", ["x"])))
    assert job.failed == 1 and "couldn't be copied" in job.message
    assert "[Image 2 could not be copied]" in fake.pages[job.page_id]["paragraphs"]


async def test_rebuild_batches_many_images(tmp_path):
    fake = FakeOneNote(copy_supported=False)
    many = "".join(f'<img src="{GRAPH}/me/onenote/resources/s{i}/$value" alt="slide {i}">' for i in range(20))
    original_content = fake.page_content
    big = fake.add_page("Big", [])

    async def content(page_id, include_ids=False):
        if page_id == big:
            return f"<html><head><title>Big</title></head><body><div>{many}<p>notes</p></div></body></html>"
        return await original_content(page_id, include_ids)

    fake.page_content = content
    ed = Editor(fake, EditState(tmp_path))
    job = await finish(ed, ed.start_copy(big))
    batches = [cmds for pid, cmds, _ in fake.updates if pid == job.page_id and cmds[0]["target"] != "title"]
    assert [len(b) for b in batches] == [8, 8, 4] and job.total == 20 and job.done == 20
    assert not find_placeholders(await fake.page_content(job.page_id, include_ids=True))


async def test_unfinished_copy_is_resumed_after_restart(tmp_path):
    fake = FakeOneNote(copy_supported=False)
    state = EditState(tmp_path)
    ed = Editor(fake, state)
    orig = fake.add_page("L2", ["notes"])

    async def broken(*a, **k):
        raise GraphError("connection dropped", None)

    from onenote_mcp.graph import GraphError

    real_get_bytes, fake.get_bytes = fake.get_bytes, broken
    real_update = fake.update_page

    async def no_fill(page_id, commands, parts):  # the network dies before any upload
        if commands[0]["target"] != "title":
            raise GraphError("connection dropped", None)
        await real_update(page_id, commands, parts)

    fake.update_page = no_fill
    job = await ed.wait(ed.start_copy(orig), 5)
    copy = job.page_id
    assert find_placeholders(await fake.page_content(copy, include_ids=True))

    # "Restart": a fresh editor sees the saved job as unfinished...
    fake.get_bytes, fake.update_page = real_get_bytes, real_update
    ed2 = Editor(fake, EditState(tmp_path))
    assert job.status == "failed" and "call copy_page again" in job.summary()
    assert ed2.jobs[job.id].status == "failed"
    # ...and an edit resumes the fill before touching the page.
    prep = await ed2.prepare(orig, None)
    resumed = ed2.latest_job_for_page(copy)
    assert resumed.id != job.id and resumed.status == "done" and "Finished copying" in resumed.message
    assert prep.page_id == copy
    assert not find_placeholders(await fake.page_content(copy, include_ids=True))
    assert len(fake.pages) == 2


async def test_edit_waits_while_copy_is_filling(tmp_path, monkeypatch):
    monkeypatch.setattr("onenote_mcp.editing.AUTO_COPY_WAIT", 0.05)
    fake = FakeOneNote(copy_supported=False)
    real_get_bytes = fake.get_bytes

    async def slow(url):
        await asyncio.sleep(0.3)
        return await real_get_bytes(url)

    fake.get_bytes = slow
    ed = Editor(fake, EditState(tmp_path))
    orig = fake.add_page("L2", ["x"])
    with pytest.raises(EditError, match="still running"):
        await ed.prepare(orig, None)
    job = ed.running_job_for(orig)
    with pytest.raises(EditError, match="still being copied in"):
        await ed.prepare(job.page_id, None)
    await finish(ed, job)


async def test_job_survives_cancelled_waiter_and_dedupes(tmp_path):
    fake = FakeOneNote(copy_delay=0.3)
    ed = Editor(fake, EditState(tmp_path))
    orig = fake.add_page("L2", ["x"])
    job = ed.start_copy(orig)
    waiter = asyncio.create_task(ed.wait(job, 10))
    await asyncio.sleep(0.05)
    waiter.cancel()  # like the user pressing Stop
    assert ed.start_copy(orig) is job  # a retry joins the running job
    assert "still running" in job.summary() and job.id in job.summary()
    await finish(ed, job)
    assert len(fake.pages) == 2


async def test_slow_auto_copy_asks_to_come_back(tmp_path, monkeypatch):
    monkeypatch.setattr("onenote_mcp.editing.AUTO_COPY_WAIT", 0.05)
    fake = FakeOneNote(copy_delay=0.3)
    ed = Editor(fake, EditState(tmp_path))
    orig = fake.add_page("L2", ["x"])
    with pytest.raises(EditError, match="get_copy_status"):
        await ed.prepare(orig, None)
    await finish(ed, ed.running_job_for(orig))
    prep = await ed.prepare(orig, None)
    assert prep.page_id != orig and len(fake.pages) == 2


async def test_timeout_falls_back_with_note(tmp_path):
    from onenote_mcp.graph import GraphError

    fake = FakeOneNote()

    async def slow(page_id, section_id, timeout=90.0, on_wait=None):
        raise GraphError("OneNote's copy didn't finish within 90 seconds.", 408)

    fake.copy_page = slow
    ed = Editor(fake, EditState(tmp_path))
    job = await finish(ed, ed.start_copy(fake.add_page("L2", ["x"])))
    assert job.method == "rebuilt" and "second copy" in job.message


async def test_explicit_copy_becomes_working_copy_once(setup):
    fake, state, ed = setup
    orig = fake.add_page("L2", ["notes"])
    first = await finish(ed, ed.start_copy(orig))
    assert fake.pages[first.page_id]["title"] == "L2 (copy)" and "working copy" in first.message
    assert (await ed.prepare(orig, None)).page_id == first.page_id

    second = await finish(ed, ed.start_copy(orig, "sec2", "Revision"))
    assert fake.pages[second.page_id] == {"title": "Revision", "section": "sec2", "paragraphs": ["notes"]}
    assert state.copy_of(orig) == first.page_id and second.page_id in state.created


async def test_created_pages_are_edited_directly(setup):
    fake, state, ed = setup
    pid = fake.add_page("Mine", ["x"])
    state.add_created(pid)
    assert (await ed.prepare(pid, None)).page_id == pid
    assert len(fake.pages) == 1


async def test_restore_rebuilds_from_snapshot(setup):
    fake, state, ed = setup
    orig = fake.add_page("Notes", ["good"])
    copy = (await ed.prepare(orig, None)).page_id
    fake.pages[copy]["paragraphs"] = ["broken"]
    job = await finish(ed, ed.start_restore(orig))
    restored = state.copy_of(orig)
    assert restored == job.page_id and restored not in (orig, copy)
    assert fake.pages[restored]["paragraphs"][0] == "good"
    assert not find_placeholders(await fake.page_content(restored, include_ids=True))
    assert fake.pages[restored]["title"] == "Notes" + COPY_SUFFIX
    assert copy in job.message and "delete it" in job.message
    with pytest.raises(EditError):
        ed.start_restore(fake.add_page("Untouched", []))
