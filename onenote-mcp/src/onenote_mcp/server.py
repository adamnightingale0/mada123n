"""The OneNote MCP server (read-only, stdio)."""

from __future__ import annotations

import asyncio
import logging
from typing import Annotated, Awaitable, Callable, Literal, TypeVar

import anyio
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import ToolAnnotations
from pydantic import Field

from .auth import AuthError, Authenticator, SignInRequired
from .compose import ContentError, Parts, attachment_tag, image_tag, markdown_to_html, page_html
from .config import ConfigError, Flags, app_dir, load_flags, load_settings
from .convert import render_page, render_structure
from .editing import EditError, Editor
from .graph import GraphClient, GraphError
from .state import EditState

log = logging.getLogger(__name__)

T = TypeVar("T")

READ_ONLY = ToolAnnotations(readOnlyHint=True, destructiveHint=False, openWorldHint=True)
ADDS = ToolAnnotations(readOnlyHint=False, destructiveHint=False, idempotentHint=False, openWorldHint=True)
CHANGES = ToolAnnotations(readOnlyHint=False, destructiveHint=True, idempotentHint=False, openWorldHint=True)

INSTRUCTIONS = """\
Read-only access to the user's Microsoft OneNote notebooks.
Navigate top-down: list_notebooks -> list_sections -> list_pages -> get_page_content,
or use search_pages to find pages by title. OneNote's API is slow and throttles, so
fetch only the pages you need rather than reading many pages in bulk.
If a tool reports that sign-in is needed, show the user the URL and code it gives
and ask them to say when they have finished signing in, then retry."""

WRITE_INSTRUCTIONS = """

Editing is enabled. Content is written as Markdown (headings, bold/italic, lists,
"- [ ]" to-dos, tables, links, ![alt](local path or URL) images). For formatting Markdown
can't express, inline HTML is allowed: <u>, <sub>, <sup>, and <span style="..."> with
color, background-color (highlight), font-size, font-family, font-weight, font-style or
text-decoration. Other HTML is removed.
- Original pages are never edited. To work on an existing page, call copy_page first
  and edit the copy's ID; this copies everything, including slide printouts, attached
  files and ink. (Editing an original's ID also works: the server makes the copy itself,
  titled "<title> [edited]", and edits that.) Tell the user the copy's name.
- Never rebuild an existing page with create_page; it can't carry over images,
  printouts or attachments. Use copy_page.
- Copying a large page (e.g. lecture slides) runs in the background. If copy_page says the
  job is still running, call get_copy_status with its job_id until it finishes; tell the
  user what it reports. Don't call copy_page again for the same page meanwhile.
- To change specific text, call get_page_structure first and use an element ID as
  target_id. Prefer replacing individual elements over rewriting the whole page.
- Before each edit the server saves one snapshot of the page; restore_previous_version
  undoes the most recent edit only.
- Edits appear in OneNote after it syncs, usually within a minute."""

Position = Literal["end", "start", "before", "after"]
PageId = Annotated[
    str, Field(description="Page ID. If it's an original page, the edit goes to its working copy.")
]
TargetId = Annotated[
    str | None, Field(description="Element ID from get_page_structure (needed for before/after)")
]
Pos = Annotated[
    Position,
    Field(description="end/start of the page (or inside target_id), or before/after target_id"),
]
Content = Annotated[str, Field(min_length=1, description="Markdown content")]


def _placement(position: str, target_id: str | None, html: str) -> dict:
    if position in ("before", "after"):
        if not target_id:
            raise ToolError(f"position={position!r} needs a target_id from get_page_structure.")
        return {"target": target_id, "action": "insert", "position": position, "content": html}
    action = "append" if position == "end" else "prepend"
    return {"target": target_id or "body", "action": action, "content": html}


class OneNoteService:
    """Wires auth and Graph together and turns failures into ToolErrors."""

    def __init__(
        self,
        auth_factory: Callable[[], Authenticator] | None = None,
        graph_factory: Callable[["OneNoteService"], GraphClient] | None = None,
        state_factory: Callable[[], EditState] | None = None,
    ):
        self._auth_factory = auth_factory or (lambda: Authenticator(load_settings()))
        self._graph_factory = graph_factory or (lambda svc: GraphClient(svc.token))
        self._state_factory = state_factory or (lambda: EditState(app_dir()))
        self._auth: Authenticator | None = None
        self._graph: GraphClient | None = None
        self._editor: Editor | None = None

    @property
    def editor(self) -> Editor:
        if self._editor is None:
            self._editor = Editor(self.graph, self._state_factory())
        return self._editor

    @property
    def auth(self) -> Authenticator:
        if self._auth is None:
            try:
                self._auth = self._auth_factory()
            except ConfigError as exc:
                raise ToolError(f"OneNote server is not configured: {exc}") from exc
        return self._auth

    @property
    def graph(self) -> GraphClient:
        if self._graph is None:
            self._graph = self._graph_factory(self)
        return self._graph

    async def token(self, force_refresh: bool = False) -> str:
        auth = self.auth
        try:
            return await anyio.to_thread.run_sync(lambda: auth.get_token(force_refresh))
        except SignInRequired:
            pass
        previous = auth.take_last_error()
        try:
            prompt = await anyio.to_thread.run_sync(auth.start_background_login)
        except AuthError as exc:
            raise ToolError(str(exc)) from exc
        lead = f"The last sign-in attempt failed ({previous}). " if previous else ""
        raise ToolError(
            f"{lead}Sign-in to Microsoft is needed before OneNote can be read. "
            f"Ask the user to open {prompt.verification_uri} in a browser and enter the code "
            f"{prompt.user_code} (valid for about {prompt.expires_in // 60} minutes), sign in "
            "with their Microsoft account and accept the OneNote permission. "
            "Once they say they're done, call this tool again."
        )

    async def call(self, fn: Callable[[GraphClient], Awaitable[T]]) -> T:
        try:
            return await fn(self.graph)
        except (GraphError, EditError, ContentError) as exc:
            raise ToolError(str(exc)) from exc

    async def edit(
        self,
        page_id: str,
        target_id: str | None,
        build: Callable[[str | None, Parts], list[dict]],
        done: str,
    ) -> str:
        """Apply commands to the page's working copy (made on first edit)."""

        async def run(graph: GraphClient) -> str:
            parts = Parts()
            prep = await self.editor.prepare(page_id, target_id)
            commands = build(prep.target, parts)
            await graph.update_page(prep.page_id, commands, parts.items)
            return "\n".join([*prep.notes, f"{done} (page id: {prep.page_id})."])

        return await self.call(run)


def _link(page: dict) -> str:
    href = ((page.get("links") or {}).get("oneNoteClientUrl") or {}).get("href")
    return f"\nOpen in OneNote: {href}" if href else ""


# How long copy/restore tools wait for their background job before reporting progress.
TOOL_WAIT = 25.0


async def _wait_and_report(g: GraphClient, job, seconds: float) -> str:
    if job.task and not job.task.done() and seconds > 0:
        await asyncio.wait({job.task}, timeout=seconds)
    return await _job_result(g, job)


async def _job_result(g: GraphClient, job) -> str:
    if job.status == "failed":
        raise ToolError(job.summary())
    out = job.summary()
    if job.status == "done" and job.page_id:
        try:
            out += _link(await g.page(job.page_id))
        except GraphError:
            pass
    return out


def _fmt_date(value: str | None) -> str:
    return value[:10] if value else "?"


def _flags() -> Flags:
    try:
        return load_flags()
    except ConfigError as exc:
        log.warning("Ignoring config for write flags: %s", exc)
        return Flags()


def build_server(service: OneNoteService | None = None, flags: Flags | None = None) -> MCPServer:
    svc = service or OneNoteService()
    flags = flags or _flags()
    instructions = INSTRUCTIONS + (WRITE_INSTRUCTIONS if flags.allow_write else "")
    mcp = MCPServer("onenote", instructions=instructions)

    @mcp.tool(annotations=READ_ONLY)
    async def list_notebooks() -> str:
        """List the user's OneNote notebooks with their IDs."""
        notebooks = await svc.call(lambda g: g.notebooks())
        if not notebooks:
            return "No notebooks found for this account."
        lines = [f"{len(notebooks)} notebook(s):"]
        for nb in notebooks:
            shared = ", shared" if nb.get("isShared") else ""
            lines.append(
                f"- {nb.get('displayName')} (id: {nb['id']}; modified "
                f"{_fmt_date(nb.get('lastModifiedDateTime'))}{shared})"
            )
        return "\n".join(lines)

    @mcp.tool(annotations=READ_ONLY)
    async def list_sections(
        notebook_id: Annotated[str, Field(description="Notebook ID from list_notebooks")],
    ) -> str:
        """List all sections in a notebook, including sections inside section groups."""
        sections = await svc.call(lambda g: g.notebook_sections(notebook_id))
        if not sections:
            return "This notebook has no sections."
        lines = [f"{len(sections)} section(s):"]
        for s in sections:
            lines.append(
                f"- {s.get('path') or s.get('displayName')} (id: {s['id']}; modified "
                f"{_fmt_date(s.get('lastModifiedDateTime'))})"
            )
        return "\n".join(lines)

    @mcp.tool(annotations=READ_ONLY)
    async def list_pages(
        section_id: Annotated[str, Field(description="Section ID from list_sections")],
        limit: Annotated[int, Field(ge=1, le=500, description="Maximum pages to return")] = 100,
    ) -> str:
        """List pages in a section (titles, IDs, dates), most recently modified first."""
        pages, more = await svc.call(lambda g: g.section_pages(section_id, limit))
        if not pages:
            return "This section has no pages."
        state = svc.editor.state if flags.allow_write else None
        lines = [f"{len(pages)} page(s){' (more exist; raise limit to see them)' if more else ''}:"]
        for p in pages:
            role = ""
            if state:
                if (orig := state.original_of(p["id"])):
                    role = f"; working copy of {orig}"
                elif (copy := state.copy_of(p["id"])):
                    role = f"; original, edits go to working copy {copy}"
                elif p["id"] in state.created:
                    role = "; created by Claude"
            lines.append(
                f"- {p.get('title') or '(untitled)'} (id: {p['id']}; created "
                f"{_fmt_date(p.get('createdDateTime'))}, modified "
                f"{_fmt_date(p.get('lastModifiedDateTime'))}{role})"
            )
        return "\n".join(lines)

    @mcp.tool(annotations=READ_ONLY)
    async def get_page_content(
        page_id: Annotated[str, Field(description="Page ID from list_pages or search_pages")],
        max_chars: Annotated[
            int, Field(ge=1000, le=200_000, description="Truncate the page text after this many characters")
        ] = 40_000,
        image_text: Annotated[
            bool,
            Field(description="Include text recognised in images (e.g. lecture slides), listed "
                              "after the page's typed text. Set false for just the typed notes."),
        ] = True,
    ) -> str:
        """Read one page as Markdown. Typed text comes first; images appear as numbered
        placeholders and attachments/printouts as [Attachment: name]. Long text recognised in
        images (slide OCR) is listed at the end. Handwriting/ink is not included."""
        html = await svc.call(lambda g: g.page_content(page_id))
        return render_page(html, page_id, max_chars, image_text)

    @mcp.tool(annotations=READ_ONLY)
    async def search_pages(
        query: Annotated[str, Field(min_length=1, description="Words to look for in page titles")],
        limit: Annotated[int, Field(ge=1, le=100, description="Maximum matches to return")] = 25,
    ) -> str:
        """Find pages whose title contains all the given words (case-insensitive).
        Searches titles only, not page text, across the most recently modified pages."""
        scan = 1000
        pages, more = await svc.call(lambda g: g.all_pages(scan))
        terms = query.lower().split()
        hits = [p for p in pages if all(t in (p.get("title") or "").lower() for t in terms)]
        if not hits:
            note = f" (searched the {scan} most recently modified pages)" if more else ""
            return f"No page titles match {query!r}{note}."
        lines = [f"{min(len(hits), limit)} match(es) for {query!r}:"]
        for p in hits[:limit]:
            nb = (p.get("parentNotebook") or {}).get("displayName", "?")
            sec = (p.get("parentSection") or {}).get("displayName", "?")
            lines.append(
                f"- {p.get('title') or '(untitled)'} — {nb} / {sec} (id: {p['id']}; modified "
                f"{_fmt_date(p.get('lastModifiedDateTime'))})"
            )
        if len(hits) > limit:
            lines.append(f"...and {len(hits) - limit} more; raise limit or refine the query.")
        if more:
            lines.append(f"(Only the {scan} most recently modified pages were searched.)")
        return "\n".join(lines)

    @mcp.tool(annotations=READ_ONLY)
    async def get_page_structure(
        page_id: Annotated[str, Field(description="Page ID")],
    ) -> str:
        """List a page's paragraphs, headings, lists, tables and images with their element
        IDs, for use as target_id in edits."""
        html = await svc.call(lambda g: g.page_content(page_id, include_ids=True))
        out = render_structure(html, page_id)
        if flags.allow_write and (copy := svc.editor.state.copy_of(page_id)):
            out += (
                f"\n\nNote: edits to this page go to its working copy (id: {copy}). Call "
                "get_page_structure on that ID to see the current, edited version."
            )
        return out

    if flags.allow_write:
        _register_write_tools(mcp, svc, flags)
    return mcp


def _register_write_tools(mcp: MCPServer, svc: OneNoteService, flags: Flags) -> None:
    @mcp.tool(annotations=ADDS)
    async def create_section(
        notebook_id: Annotated[str, Field(description="Notebook ID")],
        name: Annotated[str, Field(min_length=1, max_length=50)],
    ) -> str:
        """Create a new section in a notebook."""
        section = await svc.call(lambda g: g.create_section(notebook_id, name))
        return f"Created section '{section.get('displayName', name)}' (id: {section['id']})."

    @mcp.tool(annotations=ADDS)
    async def create_page(
        section_id: Annotated[str, Field(description="Section ID")],
        title: Annotated[str, Field(min_length=1)],
        content: Annotated[str, Field(description="Markdown content")] = "",
    ) -> str:
        """Create a new page from Markdown. New pages can be edited directly (no copy is made)."""

        async def run(g: GraphClient) -> str:
            parts = Parts()
            body = markdown_to_html(content, parts) if content.strip() else ""
            page = await g.create_page(section_id, page_html(title, body), parts.items)
            svc.editor.state.add_created(page["id"])
            return f"Created page '{title}' (id: {page['id']}).{_link(page)}"

        return await svc.call(run)

    @mcp.tool(annotations=ADDS)
    async def copy_page(
        page_id: Annotated[str, Field(description="Page to copy")],
        target_section_id: Annotated[
            str | None, Field(description="Section for the copy (default: same section)")
        ] = None,
        new_title: Annotated[str | None, Field(description="Default: '<title> (copy)'")] = None,
        allow_duplicate: Annotated[
            bool, Field(description="Make another copy even if this page already has a working copy")
        ] = False,
    ) -> str:
        """Duplicate a page, including slide printouts, images, attached files and ink, without
        changing the original. Use this before editing an existing page: the first copy of a
        page becomes its working copy. Large pages are copied in the background: if the result
        says the job is still running, follow it with get_copy_status instead of retrying.
        Calling this again for the same page returns the running job or the existing copy."""

        async def run(g: GraphClient) -> str:
            ed = svc.editor
            existing = ed.state.copy_of(page_id)
            if existing and not allow_duplicate and not ed.running_job_for(page_id):
                try:
                    meta = await g.page(existing)
                    return (
                        f"This page already has a working copy: '{meta.get('title')}' (id: {existing}). "
                        "Edit that, or pass allow_duplicate=true to make another copy." + _link(meta)
                    )
                except GraphError as exc:
                    if exc.status != 404:
                        raise
                    ed.state.forget(existing)
            job = await ed.wait(ed.start_copy(page_id, target_section_id, new_title), TOOL_WAIT)
            return await _job_result(g, job)

        return await svc.call(run)

    @mcp.tool(annotations=READ_ONLY)
    async def get_copy_status(
        job_id: Annotated[str, Field(description="Job ID from copy_page or restore_previous_version")],
        wait_seconds: Annotated[
            int, Field(ge=0, le=50, description="Wait up to this long for the job to finish")
        ] = 20,
    ) -> str:
        """Check a background copy or restore job: progress while it runs, then the new page."""
        job = svc.editor.jobs.get(job_id)
        if job is None:
            raise ToolError(
                f"No job {job_id!r}. The server may have restarted (jobs stop when Claude Desktop "
                "closes); check list_pages for the copy before starting another."
            )
        return await svc.call(lambda g: _wait_and_report(g, job, wait_seconds))

    @mcp.tool(annotations=ADDS)
    async def add_content(page_id: PageId, content: Content, position: Pos = "end",
                          target_id: TargetId = None) -> str:
        """Add Markdown content to a page: at the end or start, or before/after an element.
        Never changes an original page: if page_id is an original, the edit goes to its
        working copy (made on first edit as '<title> [edited]' in the same section). The
        result gives the ID of the page that was changed."""
        return await svc.edit(
            page_id, target_id,
            lambda t, parts: [_placement(position, t, markdown_to_html(content, parts))],
            f"Added content ({position})",
        )

    @mcp.tool(annotations=CHANGES)
    async def replace_content(
        page_id: PageId,
        target_id: Annotated[str, Field(description="Element ID from get_page_structure")],
        content: Content,
    ) -> str:
        """Replace one element (paragraph, heading, list, table, image or outline) with new
        Markdown content. Use this to rewrite or reformat existing text.
        Never changes an original page: if page_id is an original, the edit goes to its
        working copy (made on first edit as '<title> [edited]' in the same section). The
        result gives the ID of the page that was changed."""
        return await svc.edit(
            page_id, target_id,
            lambda t, parts: [{"target": t, "action": "replace", "content": markdown_to_html(content, parts)}],
            "Replaced the element",
        )

    @mcp.tool(annotations=CHANGES)
    async def set_page_title(page_id: PageId, title: Annotated[str, Field(min_length=1)]) -> str:
        """Change a page's title.
        Never changes an original page: if page_id is an original, the edit goes to its
        working copy (made on first edit as '<title> [edited]' in the same section). The
        result gives the ID of the page that was changed."""
        return await svc.edit(
            page_id, "title",
            lambda t, parts: [{"target": "title", "action": "replace", "content": title}],
            f"Set the title to '{title}'",
        )

    @mcp.tool(annotations=ADDS)
    async def insert_image(
        page_id: PageId,
        source: Annotated[str, Field(description="Local file path (png/jpg/gif/bmp/tiff) or http(s) URL")],
        alt_text: str = "",
        width: Annotated[int | None, Field(ge=16, le=2000, description="Display width in pixels")] = None,
        position: Pos = "end",
        target_id: TargetId = None,
    ) -> str:
        """Insert an image from a file on this computer or a public URL.
        Never changes an original page: if page_id is an original, the edit goes to its
        working copy (made on first edit as '<title> [edited]' in the same section). The
        result gives the ID of the page that was changed."""
        return await svc.edit(
            page_id, target_id,
            lambda t, parts: [_placement(position, t, image_tag(source, parts, alt_text, width))],
            "Inserted the image",
        )

    @mcp.tool(annotations=ADDS)
    async def attach_file(
        page_id: PageId,
        path: Annotated[str, Field(description="Local file path (max 25 MB)")],
        position: Pos = "end",
        target_id: TargetId = None,
    ) -> str:
        """Attach a file from this computer to a page (shown as a file icon in OneNote).
        Never changes an original page: if page_id is an original, the edit goes to its
        working copy (made on first edit as '<title> [edited]' in the same section). The
        result gives the ID of the page that was changed."""
        return await svc.edit(
            page_id, target_id,
            lambda t, parts: [_placement(position, t, attachment_tag(path, parts))],
            "Attached the file",
        )

    @mcp.tool(annotations=ADDS)
    async def restore_previous_version(page_id: PageId) -> str:
        """Undo the most recent edit by rebuilding the page from the snapshot taken just
        before it, as a new page in the same section. Runs in the background like copy_page;
        follow it with get_copy_status if it's still running."""
        async def run(g: GraphClient) -> str:
            job = await svc.editor.wait(svc.editor.start_restore(page_id), TOOL_WAIT)
            return await _job_result(g, job)

        return await svc.call(run)

    if flags.allow_delete:

        @mcp.tool(annotations=CHANGES)
        async def delete_page(page_id: Annotated[str, Field(description="Page ID to delete")]) -> str:
            """Permanently delete a page. Only call this when the user explicitly asks to
            delete that page. OneNote may keep it in Deleted Notes for a while."""

            async def run(g: GraphClient) -> str:
                await g.delete_page(page_id)
                svc.editor.state.forget(page_id)
                return f"Deleted page {page_id}."

            return await svc.call(run)
