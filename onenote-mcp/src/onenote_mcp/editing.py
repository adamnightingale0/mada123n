"""Copy-on-first-edit, target resolution, snapshots and restore."""

from __future__ import annotations

import logging
from dataclasses import dataclass

from bs4 import BeautifulSoup

from .compose import Parts
from .convert import page_structure
from .graph import GRAPH_BASE, GraphClient, GraphError
from .state import EditState

log = logging.getLogger(__name__)

COPY_SUFFIX = " [edited]"
# Graph statuses for which a server-side copy is worth retrying as an HTML copy.
COPY_FALLBACK_STATUSES = {400, 403, 405, 409, 501}


class EditError(Exception):
    """An edit can't go ahead. The message is for the user/model."""


@dataclass
class Prepared:
    page_id: str  # the page the edit is applied to
    target: str | None
    notes: list[str]


def _is_graph_url(url: str | None) -> bool:
    return bool(url) and url.startswith(("https://graph.microsoft.com/", GRAPH_BASE))


class Editor:
    def __init__(self, graph: GraphClient, state: EditState):
        self.graph = graph
        self.state = state

    # -- working copies ----------------------------------------------------

    async def working_page(self, page_id: str) -> tuple[str, str | None, list[str]]:
        """Return (working page ID, original ID or None, notes), copying on first edit."""
        if self.state.is_working_page(page_id):
            return page_id, self.state.original_of(page_id), []
        existing = self.state.copy_of(page_id)
        if existing:
            try:
                await self.graph.page(existing)
                return existing, page_id, [f"Editing the existing working copy (id: {existing})."]
            except GraphError as exc:
                if exc.status != 404:
                    raise
                self.state.forget(existing)
                notes = ["The previous working copy was deleted, so a new one was made."]
        else:
            notes = []

        meta = await self.graph.page(page_id)
        section_id = (meta.get("parentSection") or {}).get("id")
        if not section_id:
            raise EditError("Couldn't find the section this page is in.")
        title = meta.get("title") or "Untitled"
        try:
            copy_id = await self.graph.copy_page(page_id, section_id)
            fidelity = "an exact copy"
        except GraphError as exc:
            if exc.status not in COPY_FALLBACK_STATUSES:
                raise
            log.warning("copyToSection failed (%s); falling back to an HTML copy", exc)
            html = await self.graph.page_content(page_id)
            copy_id = await self.create_from_html(section_id, html, title)
            fidelity = (
                "a copy rebuilt from the page's HTML (Microsoft's exact-copy feature wasn't "
                "available, so handwriting/ink and drawings were not copied)"
            )
        # Record the link first, so a failure below can't lead to a second copy on retry.
        self.state.set_copy(page_id, copy_id)
        try:
            await self.graph.update_page(
                copy_id, [{"target": "title", "action": "replace", "content": title + COPY_SUFFIX}], []
            )
        except GraphError as exc:
            log.warning("Could not rename working copy: %s", exc)
            notes.append(f"(The copy couldn't be renamed, so it has the same title as the original: {exc})")
        notes.append(
            f"First edit of this page: made {fidelity} named '{title}{COPY_SUFFIX}' in the same "
            f"section (id: {copy_id}). The original page is unchanged; all edits go to the copy."
        )
        return copy_id, page_id, notes

    async def create_from_html(self, section_id: str, html: str, title: str) -> str:
        """Create a page from existing page HTML, re-uploading its images and files."""
        soup = BeautifulSoup(html, "html.parser")
        parts = Parts()
        for el in soup.find_all(True):  # IDs belong to the source page
            el.attrs.pop("id", None)
            el.attrs.pop("data-id", None)
        for img in soup.find_all("img"):
            src = img.get("data-fullres-src") or img.get("src")
            if _is_graph_url(src):
                try:
                    data, ctype = await self.graph.get_bytes(src)
                    img["src"] = f"name:{parts.add(data, ctype.split(';')[0], 'img')}"
                except GraphError:
                    img.replace_with(f"[Image could not be copied: {img.get('alt', '')}]")
                    continue
            for attr in ("data-fullres-src", "data-fullres-src-type", "data-src-type"):
                img.attrs.pop(attr, None)
        for obj in soup.find_all("object"):
            src = obj.get("data")
            if _is_graph_url(src):
                try:
                    data, ctype = await self.graph.get_bytes(src)
                    obj["data"] = f"name:{parts.add(data, obj.get('type') or ctype, 'file')}"
                except GraphError:
                    obj.replace_with(f"[Attachment could not be copied: {obj.get('data-attachment', '')}]")
        if soup.title is not None:
            soup.title.string = title
        page = await self.graph.create_page(section_id, str(soup), parts.items)
        return page["id"]

    # -- targets -----------------------------------------------------------

    async def prepare(self, page_id: str, target: str | None) -> Prepared:
        """Find the working page, snapshot it, and map ``target`` onto it."""
        working, original, notes = await self.working_page(page_id)
        html = await self.graph.page_content(working, include_ids=True)
        self.state.save_snapshot(working, html)

        if target in (None, "title", "body"):
            return Prepared(working, target, notes)
        structure = page_structure(html)
        if any(e.target == target for e in structure):
            return Prepared(working, target, notes)
        if original:
            # The ID came from the original page; find the same element in the copy.
            orig = page_structure(await self.graph.page_content(original, include_ids=True))
            idx = next((i for i, e in enumerate(orig) if e.target == target), None)
            if idx is not None and idx < len(structure):
                cand = structure[idx]
                if cand.tag == orig[idx].tag and cand.preview == orig[idx].preview:
                    return Prepared(working, cand.target, notes)
        raise EditError(
            f"Element {target!r} isn't on the working page (id: {working}). Call "
            "get_page_structure with that page ID to get current element IDs."
            + (" " + " ".join(notes) if notes else "")
        )

    # -- restore -----------------------------------------------------------

    async def restore(self, page_id: str) -> str:
        """Rebuild the working page from its snapshot as a new page. Returns a message."""
        working = page_id if self.state.is_working_page(page_id) else self.state.copy_of(page_id)
        if not working:
            raise EditError("This page hasn't been edited by Claude, so there's no snapshot to restore.")
        html = self.state.load_snapshot(working)
        if html is None:
            raise EditError("No snapshot is saved for this page.")
        meta = await self.graph.page(working)
        section_id = (meta.get("parentSection") or {}).get("id")
        if not section_id:
            raise EditError("Couldn't find the section this page is in.")
        title = meta.get("title") or "Untitled"
        new_id = await self.create_from_html(section_id, html, title)

        original = self.state.original_of(working)
        self.state.forget(working)
        if original:
            self.state.set_copy(original, new_id)
        else:
            self.state.add_created(new_id)
        self.state.save_snapshot(new_id, html)
        return (
            f"Restored '{title}' as it was before the last edit, as a new page (id: {new_id}) in "
            f"the same section. Further edits go to the restored page. The page with the bad edit "
            f"(id: {working}) is still there; delete it in OneNote when you've checked the restore."
        )
