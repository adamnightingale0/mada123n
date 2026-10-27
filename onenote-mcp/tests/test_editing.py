import pytest

from onenote_mcp.editing import COPY_SUFFIX, EditError, Editor
from onenote_mcp.state import EditState

from .fake_onenote import FakeOneNote


@pytest.fixture
def setup(tmp_path):
    fake = FakeOneNote()
    state = EditState(tmp_path)
    return fake, state, Editor(fake, state)


async def test_first_edit_copies_and_later_edits_reuse(setup, tmp_path):
    fake, state, ed = setup
    orig = fake.add_page("Notes", ["one", "two"])
    prep = await ed.prepare(orig, None)
    copy = prep.page_id
    assert copy != orig and fake.pages[copy]["title"] == "Notes" + COPY_SUFFIX
    assert fake.pages[orig]["title"] == "Notes"  # original untouched
    assert "First edit" in prep.notes[0]

    again = await ed.prepare(orig, None)  # via the original's ID
    assert again.page_id == copy and "existing working copy" in again.notes[0]
    direct = await ed.prepare(copy, None)  # via the copy's ID
    assert direct.page_id == copy and direct.notes == []
    assert len(fake.pages) == 2

    # State survives a restart.
    assert EditState(tmp_path).copy_of(orig) == copy


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
    fake, state, ed = setup
    orig = fake.add_page("Notes", ["alpha", "beta"])
    prep = await ed.prepare(orig, f"p:{{{orig}}}{{2}}")
    assert prep.target == f"p:{{{prep.page_id}}}{{2}}"


async def test_unknown_target_is_reported(setup):
    fake, _, ed = setup
    orig = fake.add_page("Notes", ["alpha"])
    with pytest.raises(EditError, match="get_page_structure"):
        await ed.prepare(orig, "p:{nope}{9}")


async def test_deleted_copy_is_recreated(setup):
    fake, state, ed = setup
    orig = fake.add_page("Notes", ["a"])
    first = (await ed.prepare(orig, None)).page_id
    del fake.pages[first]
    prep = await ed.prepare(orig, None)
    assert prep.page_id != first and "previous working copy was deleted" in prep.notes[0]


async def test_falls_back_to_html_copy(tmp_path):
    fake = FakeOneNote(copy_supported=False)
    ed = Editor(fake, EditState(tmp_path))
    orig = fake.add_page("Notes", ["a"])
    prep = await ed.prepare(orig, None)
    assert "ink" in prep.notes[0]
    section, html, parts = fake.created[0]
    assert section == "sec1" and 'src="name:img1"' in html and parts[0].data == b"\x89PNG"
    assert ' id="' not in html


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
    msg = await ed.restore(orig)
    restored = state.copy_of(orig)
    assert restored not in (orig, copy)
    assert fake.pages[restored]["paragraphs"] == ["good"]
    assert fake.pages[restored]["title"] == "Notes" + COPY_SUFFIX
    assert copy in msg and "delete it" in msg
    with pytest.raises(EditError):
        await ed.restore(fake.add_page("Untouched", []))
