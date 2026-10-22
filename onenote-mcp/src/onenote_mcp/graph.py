"""Minimal async Microsoft Graph client for the OneNote read endpoints."""

from __future__ import annotations

import asyncio
import logging
from typing import Any, Awaitable, Callable
from urllib.parse import quote

import httpx

log = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# Statuses worth retrying after a pause: throttling and transient server errors.
RETRY_STATUSES = {429, 502, 503, 504}
MAX_RETRIES = 3
MAX_RETRY_WAIT = 30.0

TokenProvider = Callable[[bool], Awaitable[str]]


class GraphError(Exception):
    """A Graph request failed. The message is written for the user/model to read."""

    def __init__(self, message: str, status: int | None = None):
        super().__init__(message)
        self.status = status


def segment(item_id: str) -> str:
    """Quote an ID for use as one URL path segment (OneNote IDs contain '!')."""
    if not item_id or not item_id.strip():
        raise GraphError("An ID is required. Use the list_* tools to find valid IDs.")
    return quote(item_id.strip(), safe="!-_.~")


def _graph_message(resp: httpx.Response) -> str:
    try:
        err = resp.json().get("error", {})
    except ValueError:
        return resp.text[:300].strip()
    code = err.get("code")
    msg = err.get("message") or ""
    return f"{msg} (Graph error code {code})" if code else msg


def _retry_after(resp: httpx.Response, attempt: int) -> float:
    header = resp.headers.get("Retry-After")
    if header:
        try:
            return min(float(header), MAX_RETRY_WAIT)
        except ValueError:
            pass
    return min(2.0 ** (attempt + 1), MAX_RETRY_WAIT)


def _error_for(resp: httpx.Response) -> GraphError:
    status = resp.status_code
    detail = _graph_message(resp)
    if status == 401:
        text = (
            "Microsoft rejected the sign-in token (HTTP 401). The session may have been "
            "revoked; run `onenote-mcp logout` then sign in again."
        )
    elif status == 403:
        text = (
            "Access denied by Microsoft Graph (HTTP 403). Check that the app registration "
            "has the delegated Notes.Read permission and that you consented to it."
        )
    elif status == 404:
        text = (
            "Not found (HTTP 404). The ID may be wrong, or the item was deleted or moved. "
            "Use the list_* tools to get current IDs."
        )
    elif status == 429:
        wait = resp.headers.get("Retry-After")
        text = (
            "Microsoft Graph is throttling OneNote requests (HTTP 429)"
            + (f"; it asked to wait {wait} seconds" if wait else "")
            + ". Wait a little and retry, and avoid fetching many pages in a row."
        )
    elif status >= 500:
        text = f"Microsoft Graph had a server error (HTTP {status}). Try again shortly."
    else:
        text = f"Microsoft Graph request failed (HTTP {status})."
    if detail:
        text += f" Details: {detail}"
    return GraphError(text, status)


class GraphClient:
    def __init__(
        self,
        token_provider: TokenProvider,
        http: httpx.AsyncClient | None = None,
        sleep: Callable[[float], Awaitable[None]] = asyncio.sleep,
    ):
        self._token = token_provider
        self._http = http or httpx.AsyncClient(timeout=httpx.Timeout(60.0, connect=15.0))
        self._sleep = sleep

    async def aclose(self) -> None:
        await self._http.aclose()

    async def _request(self, url: str, params: dict[str, Any] | None = None) -> httpx.Response:
        if not url.startswith("https://"):
            url = GRAPH_BASE + url
        refreshed = False
        attempt = 0
        while True:
            token = await self._token(refreshed)
            try:
                resp = await self._http.get(
                    url, params=params, headers={"Authorization": f"Bearer {token}"}
                )
            except httpx.TimeoutException as exc:
                raise GraphError(
                    "Microsoft Graph timed out. OneNote can be slow; try again, or ask for "
                    "a smaller set of items."
                ) from exc
            except httpx.HTTPError as exc:
                raise GraphError(f"Could not reach Microsoft Graph: {exc}") from exc

            if resp.status_code == 401 and not refreshed:
                refreshed = True
                continue
            if resp.status_code in RETRY_STATUSES and attempt < MAX_RETRIES:
                wait = _retry_after(resp, attempt)
                log.warning("Graph returned %s; retrying in %.1fs", resp.status_code, wait)
                attempt += 1
                await self._sleep(wait)
                continue
            if resp.is_success:
                return resp
            raise _error_for(resp)

    async def get_json(self, url: str, params: dict[str, Any] | None = None) -> dict:
        resp = await self._request(url, params)
        try:
            return resp.json()
        except ValueError as exc:
            raise GraphError("Microsoft Graph returned a response that was not JSON.") from exc

    async def get_text(self, url: str, params: dict[str, Any] | None = None) -> str:
        return (await self._request(url, params)).text

    async def get_collection(
        self, url: str, params: dict[str, Any] | None = None, max_items: int = 500
    ) -> tuple[list[dict], bool]:
        """Fetch a collection, following @odata.nextLink.

        Returns (items, truncated) where truncated means more items exist.
        """
        items: list[dict] = []
        next_url: str | None = url
        next_params = params
        while next_url:
            data = await self.get_json(next_url, next_params)
            items.extend(data.get("value", []))
            next_url = data.get("@odata.nextLink")
            next_params = None  # nextLink already carries the query string
            if len(items) >= max_items:
                return items[:max_items], bool(next_url) or len(items) > max_items
        return items, False

    # -- OneNote endpoints -------------------------------------------------

    async def notebooks(self) -> list[dict]:
        items, _ = await self.get_collection(
            "/me/onenote/notebooks",
            {"$select": "id,displayName,lastModifiedDateTime,isShared", "$orderby": "displayName"},
        )
        return items

    async def notebook_sections(self, notebook_id: str) -> list[dict]:
        """All sections in a notebook, including those nested in section groups.

        Each section gets a ``path`` key such as "Group / Subgroup / Section".
        """
        select = {"$select": "id,displayName,lastModifiedDateTime"}
        nb = segment(notebook_id)
        top, _ = await self.get_collection(f"/me/onenote/notebooks/{nb}/sections", select)
        groups, _ = await self.get_collection(f"/me/onenote/notebooks/{nb}/sectionGroups", select)
        for s in top:
            s["path"] = s.get("displayName", "")
        out = list(top)
        for g in groups:
            out.extend(await self._group_sections(g, g.get("displayName", ""), depth=0))
        return out

    async def _group_sections(self, group: dict, prefix: str, depth: int) -> list[dict]:
        if depth > 8:  # OneNote nesting is shallow; guard against surprises
            return []
        select = {"$select": "id,displayName,lastModifiedDateTime"}
        gid = segment(group["id"])
        sections, _ = await self.get_collection(f"/me/onenote/sectionGroups/{gid}/sections", select)
        for s in sections:
            s["path"] = f"{prefix} / {s.get('displayName', '')}"
        subgroups, _ = await self.get_collection(
            f"/me/onenote/sectionGroups/{gid}/sectionGroups", select
        )
        for sg in subgroups:
            sections.extend(
                await self._group_sections(sg, f"{prefix} / {sg.get('displayName', '')}", depth + 1)
            )
        return sections

    async def section_pages(self, section_id: str, max_items: int) -> tuple[list[dict], bool]:
        return await self.get_collection(
            f"/me/onenote/sections/{segment(section_id)}/pages",
            {"$select": "id,title,createdDateTime,lastModifiedDateTime", "$top": "100"},
            max_items=max_items,
        )

    async def all_pages(self, max_items: int) -> tuple[list[dict], bool]:
        return await self.get_collection(
            "/me/onenote/pages",
            {
                "$expand": "parentSection,parentNotebook",
                "$orderby": "lastModifiedDateTime desc",
                "$top": "100",
            },
            max_items=max_items,
        )

    async def page_content(self, page_id: str) -> str:
        return await self.get_text(f"/me/onenote/pages/{segment(page_id)}/content")
