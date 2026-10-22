"""The OneNote MCP server (read-only, stdio)."""

from __future__ import annotations

import logging
from typing import Annotated, Awaitable, Callable, TypeVar

import anyio
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import ToolAnnotations
from pydantic import Field

from .auth import AuthError, Authenticator, SignInRequired
from .config import ConfigError, load_settings
from .convert import render_page
from .graph import GraphClient, GraphError

log = logging.getLogger(__name__)

T = TypeVar("T")

READ_ONLY = ToolAnnotations(readOnlyHint=True, destructiveHint=False, openWorldHint=True)

INSTRUCTIONS = """\
Read-only access to the user's Microsoft OneNote notebooks.
Navigate top-down: list_notebooks -> list_sections -> list_pages -> get_page_content,
or use search_pages to find pages by title. OneNote's API is slow and throttles, so
fetch only the pages you need rather than reading many pages in bulk.
If a tool reports that sign-in is needed, show the user the URL and code it gives
and ask them to say when they have finished signing in, then retry."""


class OneNoteService:
    """Wires auth and Graph together and turns failures into ToolErrors."""

    def __init__(
        self,
        auth_factory: Callable[[], Authenticator] | None = None,
        graph_factory: Callable[["OneNoteService"], GraphClient] | None = None,
    ):
        self._auth_factory = auth_factory or (lambda: Authenticator(load_settings()))
        self._graph_factory = graph_factory or (lambda svc: GraphClient(svc.token))
        self._auth: Authenticator | None = None
        self._graph: GraphClient | None = None

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
            "with their Microsoft account and accept the Notes.Read permission. "
            "Once they say they're done, call this tool again."
        )

    async def call(self, fn: Callable[[GraphClient], Awaitable[T]]) -> T:
        try:
            return await fn(self.graph)
        except GraphError as exc:
            raise ToolError(str(exc)) from exc


def _fmt_date(value: str | None) -> str:
    return value[:10] if value else "?"


def build_server(service: OneNoteService | None = None) -> MCPServer:
    svc = service or OneNoteService()
    mcp = MCPServer("onenote", instructions=INSTRUCTIONS)

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
        lines = [f"{len(pages)} page(s){' (more exist; raise limit to see them)' if more else ''}:"]
        for p in pages:
            lines.append(
                f"- {p.get('title') or '(untitled)'} (id: {p['id']}; created "
                f"{_fmt_date(p.get('createdDateTime'))}, modified "
                f"{_fmt_date(p.get('lastModifiedDateTime'))})"
            )
        return "\n".join(lines)

    @mcp.tool(annotations=READ_ONLY)
    async def get_page_content(
        page_id: Annotated[str, Field(description="Page ID from list_pages or search_pages")],
        max_chars: Annotated[
            int, Field(ge=1000, le=200_000, description="Truncate the page text after this many characters")
        ] = 40_000,
    ) -> str:
        """Read one page as Markdown. Images and attachments appear as placeholders;
        handwriting/ink is not included."""
        html = await svc.call(lambda g: g.page_content(page_id))
        return render_page(html, page_id, max_chars)

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

    return mcp
