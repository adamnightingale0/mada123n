"""Minimal async Microsoft Graph client for the OneNote read endpoints."""

from __future__ import annotations

import asyncio
import json
import logging
import time
from dataclasses import dataclass
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


@dataclass
class Part:
    """A binary part (image or file) sent alongside page HTML in a multipart request."""

    name: str
    data: bytes
    content_type: str


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
            "has the delegated Notes.Read permission (Notes.ReadWrite for editing) and that you "
            "consented to it; after adding a permission, run `onenote-mcp logout` and `login`."
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
    elif status == 501:
        text = "Microsoft Graph doesn't support this for your account or notebook (HTTP 501)."
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

    async def _request(
        self,
        url: str,
        params: dict[str, Any] | None = None,
        method: str = "GET",
        headers: dict[str, str] | None = None,
        **body: Any,
    ) -> httpx.Response:
        """Send a request with auth, retrying throttling and transient errors.

        Writes (POST/PATCH) are retried only on 429, which Graph sends before
        doing any work; a 5xx on a write may have half-applied, so retrying it
        could duplicate content.
        """
        if not url.startswith("https://"):
            url = GRAPH_BASE + url
        retryable = RETRY_STATUSES if method in ("GET", "DELETE") else {429}
        refreshed = False
        attempt = 0
        path = url.removeprefix(GRAPH_BASE)
        while True:
            token = await self._token(refreshed)
            started = time.monotonic()
            try:
                resp = await self._http.request(
                    method,
                    url,
                    params=params,
                    headers={**(headers or {}), "Authorization": f"Bearer {token}"},
                    **body,
                )
            except httpx.TimeoutException as exc:
                log.warning("%s %s timed out after %.1fs", method, path, time.monotonic() - started)
                raise GraphError(
                    "Microsoft Graph timed out. OneNote can be slow; try again, or ask for "
                    "a smaller set of items."
                    + (" Check the page before retrying: the change may have been applied."
                       if method != "GET" else "")
                ) from exc
            except httpx.HTTPError as exc:
                raise GraphError(f"Could not reach Microsoft Graph: {exc}") from exc

            elapsed = time.monotonic() - started
            log.log(
                logging.INFO if method != "GET" or elapsed > 3 or not resp.is_success else logging.DEBUG,
                "%s %s -> %s in %.1fs", method, path, resp.status_code, elapsed,
            )
            if resp.status_code == 401 and not refreshed:
                refreshed = True
                continue
            if resp.status_code in retryable and attempt < MAX_RETRIES:
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

    async def get_bytes(self, url: str) -> tuple[bytes, str]:
        resp = await self._request(url)
        return resp.content, resp.headers.get("Content-Type", "application/octet-stream")

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

    async def page_content(self, page_id: str, include_ids: bool = False) -> str:
        params = {"includeIDs": "true"} if include_ids else None
        return await self.get_text(f"/me/onenote/pages/{segment(page_id)}/content", params)

    async def page(self, page_id: str) -> dict:
        return await self.get_json(
            f"/me/onenote/pages/{segment(page_id)}",
            {"$select": "id,title,links", "$expand": "parentSection($select=id,displayName)"},
        )

    # -- writes ------------------------------------------------------------

    async def create_section(self, notebook_id: str, name: str) -> dict:
        resp = await self._request(
            f"/me/onenote/notebooks/{segment(notebook_id)}/sections",
            method="POST",
            json={"displayName": name},
        )
        return resp.json()

    async def create_page(self, section_id: str, html: str, parts: list[Part]) -> dict:
        url = f"/me/onenote/sections/{segment(section_id)}/pages"
        if parts:
            files = [("Presentation", (None, html.encode("utf-8"), "text/html"))]
            files += [(p.name, (p.name, p.data, p.content_type)) for p in parts]
            resp = await self._request(url, method="POST", files=files)
        else:
            resp = await self._request(
                url, method="POST", content=html.encode("utf-8"),
                headers={"Content-Type": "text/html"},
            )
        return resp.json()

    async def update_page(self, page_id: str, commands: list[dict], parts: list[Part]) -> None:
        url = f"/me/onenote/pages/{segment(page_id)}/content"
        if parts:
            files = [("Commands", (None, json.dumps(commands).encode("utf-8"), "application/json"))]
            files += [(p.name, (p.name, p.data, p.content_type)) for p in parts]
            await self._request(url, method="PATCH", files=files)
        else:
            await self._request(url, method="PATCH", json=commands)

    async def delete_page(self, page_id: str) -> None:
        await self._request(f"/me/onenote/pages/{segment(page_id)}", method="DELETE")

    async def copy_page(
        self,
        page_id: str,
        section_id: str,
        timeout: float = 90.0,
        on_wait: Callable[[float], None] | None = None,
    ) -> str:
        """Copy a page server-side (full fidelity) and return the new page's ID.

        Raises GraphError if OneNote refuses, reports failure, or hasn't finished
        within ``timeout`` seconds (status 408 in that case).
        """
        resp = await self._request(
            f"/me/onenote/pages/{segment(page_id)}/copyToSection",
            method="POST",
            json={"id": section_id},
        )
        op_url = resp.headers.get("Operation-Location")
        if not op_url:
            op_id = (resp.json() if resp.content else {}).get("id")
            if not op_id:
                raise GraphError("Microsoft Graph accepted the copy but gave no way to track it.")
            op_url = f"/me/onenote/operations/{segment(op_id)}"
        start = time.monotonic()
        waited = 0.0
        delay = 1.0
        while max(waited, time.monotonic() - start) < timeout:
            await self._sleep(delay)
            waited += delay
            if on_wait:
                on_wait(max(waited, time.monotonic() - start))
            op = await self.get_json(op_url)
            status = (op.get("status") or "").lower()
            log.info("copyToSection operation status: %s", status or op)
            if status == "completed":
                new_id = op.get("resourceId")
                if not new_id:
                    raise GraphError("The page copy finished but Graph did not return the new page ID.")
                return new_id
            if status == "failed":
                detail = (op.get("error") or {}).get("message") or "no details"
                raise GraphError(f"OneNote's copy failed: {detail}", 500)
            delay = min(delay * 1.5, 5.0)
        raise GraphError(f"OneNote's copy didn't finish within {timeout:.0f} seconds.", 408)
