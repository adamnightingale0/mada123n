"""Copy jobs, copy-on-first-edit, target resolution, snapshots and restore."""

from __future__ import annotations

import asyncio
import itertools
import logging
import time
from dataclasses import dataclass, field

from bs4 import BeautifulSoup

from .convert import page_structure
from .graph import GRAPH_BASE, GraphClient, GraphError, Part
from .state import EditState

log = logging.getLogger(__name__)

COPY_SUFFIX = " [edited]"
# Graph outcomes after which an exact copy is abandoned for a rebuilt one:
# refused (400/403/405/409/501), failed (500) or not finished in time (408).
COPY_FALLBACK_STATUSES = {400, 403, 405, 408, 409, 500, 501}
EXACT_COPY_TIMEOUT = 90.0
# How long an edit waits for its working copy before asking the model to come back.
AUTO_COPY_WAIT = 25.0
DOWNLOAD_CONCURRENCY = 6
BATCH_MAX_ITEMS = 8
BATCH_MAX_BYTES = 20 * 1024 * 1024
PLACEHOLDER_PREFIX = "onenote-mcp-pending-"


class EditError(Exception):
    """An edit can't go ahead. The message is for the user/model."""


@dataclass
class Prepared:
    page_id: str  # the page the edit is applied to
    target: str | None
    notes: list[str]


def _is_graph_url(url: str | None) -> bool:
    return bool(url) and url.startswith(("https://graph.microsoft.com/", GRAPH_BASE))


@dataclass
class _Resource:
    placeholder: str
    kind: str  # "img" or "object"
    url: str
    attrs: dict
    number: int


@dataclass
class CopyJob:
    """A page copy (or restore) running in the background.

    It keeps going if the tool call that started it is cancelled, so an
    interrupted chat never leaves the outcome unknown: ask for its status.
    """

    id: str
    kind: str  # "copy" or "restore"
    source_id: str
    make_working: bool = False
    title: str = ""
    status: str = "running"  # running / done / failed
    stage: str = "starting"
    method: str | None = None  # "exact" or "rebuilt"
    page_id: str | None = None
    total: int = 0
    done: int = 0
    failed: int = 0
    note: str = ""
    message: str = ""
    error: str = ""
    started: float = field(default_factory=time.monotonic)
    finished: float | None = None
    task: asyncio.Task | None = None

    def summary(self) -> str:
        elapsed = (self.finished or time.monotonic()) - self.started
        if self.status == "done":
            return self.message
        if self.status == "failed":
            out = f"{self.kind.capitalize()} job {self.id} failed after {elapsed:.0f}s: {self.error}"
            if self.page_id:
                out += (
                    f" A partly copied page exists (id: {self.page_id}) with {self.done - self.failed}"
                    f" of {self.total} images/files; delete it in OneNote if it isn't useful."
                )
            return out
        out = (
            f"{self.kind.capitalize()} job {self.id} is still running ({elapsed:.0f}s so far): "
            f"{self.stage}."
        )
        if self.page_id:
            out += f" The new page already exists (id: {self.page_id}) and is filling in."
        return out + (
            " It continues in the background even if this chat is interrupted. Call "
            f"get_copy_status with job_id {self.id!r} to follow it; don't start another copy."
        )


class Editor:
    def __init__(self, graph: GraphClient, state: EditState):
        self.graph = graph
        self.state = state
        self.jobs: dict[str, CopyJob] = {}
        self._job_numbers = itertools.count(1)

    # -- jobs --------------------------------------------------------------

    def _start(self, job: CopyJob, work) -> CopyJob:
        self.jobs[job.id] = job

        async def runner() -> None:
            log.info("%s job %s started for page %s", job.kind, job.id, job.source_id)
            try:
                await work(job)
                job.status = "done"
            except Exception as exc:  # reported through the job, never lost
                log.exception("%s job %s failed", job.kind, job.id)
                job.status = "failed"
                job.error = str(exc) or type(exc).__name__
            finally:
                job.finished = time.monotonic()
                log.info("%s job %s %s in %.1fs: %s", job.kind, job.id, job.status,
                         job.finished - job.started, job.message or job.error)

        job.task = asyncio.get_running_loop().create_task(runner())
        return job

    def running_job_for(self, source_id: str) -> CopyJob | None:
        return next(
            (j for j in self.jobs.values() if j.status == "running" and j.source_id == source_id), None
        )

    async def wait(self, job: CopyJob, seconds: float) -> CopyJob:
        """Wait up to ``seconds`` for a job; never cancels it."""
        if job.task and not job.task.done() and seconds > 0:
            await asyncio.wait({job.task}, timeout=seconds)
        return job

    def start_copy(
        self,
        source_id: str,
        section_id: str | None = None,
        new_title: str | None = None,
        suffix: str = " (copy)",
        make_working: bool = True,
    ) -> CopyJob:
        """Start copying a page, or return the copy job already running for it."""
        running = self.running_job_for(source_id)
        if running:
            return running
        job = CopyJob(id=f"copy-{next(self._job_numbers)}", kind="copy", source_id=source_id,
                      make_working=make_working)

        async def work(job: CopyJob) -> None:
            await self._copy(job, section_id, new_title, suffix)

        return self._start(job, work)

    async def _copy(self, job: CopyJob, section_id: str | None, new_title: str | None, suffix: str) -> None:
        meta = await self.graph.page(job.source_id)
        section_id = section_id or (meta.get("parentSection") or {}).get("id")
        if not section_id:
            raise EditError("Couldn't find the section this page is in.")
        job.title = new_title or (meta.get("title") or "Untitled") + suffix

        job.stage = "asking OneNote for an exact copy"

        def waiting(seconds: float) -> None:
            job.stage = f"waiting for OneNote's exact copy ({seconds:.0f}s)"

        try:
            job.page_id = await self.graph.copy_page(
                job.source_id, section_id, EXACT_COPY_TIMEOUT, on_wait=waiting
            )
            job.method = "exact"
            job.stage = "renaming the copy"
            try:
                await self.graph.update_page(
                    job.page_id, [{"target": "title", "action": "replace", "content": job.title}], []
                )
            except GraphError as exc:
                job.title = meta.get("title") or "Untitled"
                job.note = f"It couldn't be renamed, so it has the original's title ({exc})."
        except GraphError as exc:
            if exc.status not in COPY_FALLBACK_STATUSES:
                raise
            log.warning("Exact copy unavailable (%s); rebuilding the page instead", exc)
            job.method = "rebuilt"
            job.note = (
                f"OneNote's exact copy wasn't available ({exc}), so the page was rebuilt from its "
                "content; handwriting/ink and drawings are not included."
            )
            if exc.status == 408:
                job.note += (
                    " If a second copy with the original's title appears later, OneNote finished "
                    "its exact copy after all; delete whichever you don't want."
                )
            job.stage = "reading the page"
            html = await self.graph.page_content(job.source_id)
            await self._rebuild(job, section_id, html, job.title)

        role = self._register_copy(job)
        what = (
            "an exact copy (slide printouts, images, attachments and ink included)"
            if job.method == "exact"
            else f"a rebuilt copy with {job.done - job.failed} of {job.total} images/files"
        )
        job.message = " ".join(
            x for x in (
                f"Made {what}: '{job.title}' (id: {job.page_id}).",
                f"{job.failed} image(s)/file(s) couldn't be copied and are marked on the page."
                if job.failed else "",
                job.note,
                role,
            ) if x
        )

    def _register_copy(self, job: CopyJob) -> str:
        source = job.source_id
        if job.make_working and not self.state.is_working_page(source) and not self.state.copy_of(source):
            self.state.set_copy(source, job.page_id)
            return (
                "It is now this page's working copy: edits made with either ID go to the copy, "
                "and the original stays unchanged."
            )
        self.state.add_created(job.page_id)
        return "It is a separate page; edit it by its ID."

    # -- rebuilding a page from HTML -----------------------------------------

    async def _rebuild(self, job: CopyJob, section_id: str, html: str, title: str) -> None:
        """Create a page from existing page HTML, then upload its images/files in batches.

        The page is created first with a placeholder for each image and file, so
        it exists (and shows progress) straight away; each batch replaces its
        placeholders. Downloads run a few at a time.
        """
        soup = BeautifulSoup(html, "html.parser")
        for el in soup.find_all(True):  # IDs belong to the source page
            el.attrs.pop("id", None)
            el.attrs.pop("data-id", None)
        resources: list[_Resource] = []
        for el in soup.find_all(["img", "object"]):
            url = el.get("src") or el.get("data-fullres-src") if el.name == "img" else el.get("data")
            if not _is_graph_url(url):
                continue
            keep = ("alt", "width", "height", "style") if el.name == "img" else (
                "data-attachment", "type", "style")
            n = len(resources) + 1
            res = _Resource(f"{PLACEHOLDER_PREFIX}{n}", el.name, url,
                            {k: el[k] for k in keep if el.get(k)}, n)
            resources.append(res)
            ph = soup.new_tag("p")
            ph["data-id"] = res.placeholder
            el.replace_with(ph)
        job.total = len(resources)
        for res in resources:
            ph = soup.find("p", attrs={"data-id": res.placeholder})
            kind = "image" if res.kind == "img" else "file"
            ph.string = f"[Copying {kind} {res.number} of {job.total}…]"
        if soup.title is not None:
            soup.title.string = title

        job.stage = "creating the new page"
        page = await self.graph.create_page(section_id, str(soup), [])
        job.page_id = page["id"]

        sem = asyncio.Semaphore(DOWNLOAD_CONCURRENCY)

        async def fetch(res: _Resource):
            async with sem:
                try:
                    data, ctype = await self.graph.get_bytes(res.url)
                    return res, data, ctype.split(";")[0]
                except GraphError as exc:
                    log.warning("Couldn't download %s %d: %s", res.kind, res.number, exc)
                    return res, None, str(exc)

        for i in range(0, len(resources), BATCH_MAX_ITEMS):
            chunk = resources[i:i + BATCH_MAX_ITEMS]
            job.stage = f"copying images and files ({job.done}/{job.total})"
            fetched = await asyncio.gather(*(fetch(r) for r in chunk))
            batch: list[tuple[_Resource, bytes, str]] = []
            size = 0
            for res, data, ctype in fetched:
                if data is None:
                    await self._fill(job, [(res, None, "")])
                    continue
                if batch and (size + len(data) > BATCH_MAX_BYTES):
                    await self._fill(job, batch)
                    batch, size = [], 0
                batch.append((res, data, ctype))
                size += len(data)
            if batch:
                await self._fill(job, batch)
        job.stage = "finished"

    async def _fill(self, job: CopyJob, items: list[tuple[_Resource, bytes | None, str]]) -> None:
        """Replace placeholders with uploaded images/files; mark ones that fail."""
        commands, parts = [], []
        for res, data, ctype in items:
            if data is None:
                commands.append(self._failed_command(res))
                job.failed += 1
                continue
            name = f"{'img' if res.kind == 'img' else 'file'}{res.number}"
            parts.append(Part(name, data, ctype if res.kind == "img" else res.attrs.get("type", ctype)))
            tag = BeautifulSoup("", "html.parser").new_tag(res.kind, attrs=dict(res.attrs))
            tag["src" if res.kind == "img" else "data"] = f"name:{name}"
            commands.append({"target": f"#{res.placeholder}", "action": "replace", "content": str(tag)})
        try:
            await self.graph.update_page(job.page_id, commands, parts)
        except GraphError as exc:
            if len(items) > 1 and exc.status != 429:
                log.warning("Batch upload failed (%s); retrying items one by one", exc)
                for item in items:
                    await self._fill(job, [item])
                return
            log.warning("Upload of %s %d failed: %s", items[0][0].kind, items[0][0].number, exc)
            job.failed += sum(1 for _, data, _ in items if data is not None)
            try:
                await self.graph.update_page(job.page_id, [self._failed_command(r) for r, _, _ in items], [])
            except GraphError:
                pass
        job.done += len(items)
        job.stage = f"copying images and files ({job.done}/{job.total})"

    @staticmethod
    def _failed_command(res: _Resource) -> dict:
        kind = "Image" if res.kind == "img" else "Attached file"
        label = res.attrs.get("data-attachment", "")
        return {"target": f"#{res.placeholder}", "action": "replace",
                "content": f"<p>[{kind} {res.number} could not be copied{': ' + label if label else ''}]</p>"}

    # -- working copies ----------------------------------------------------

    async def working_page(self, page_id: str) -> tuple[str, str | None, list[str]]:
        """Return (working page ID, original ID or None, notes), copying on first edit."""
        if self.state.is_working_page(page_id):
            return page_id, self.state.original_of(page_id), []
        notes: list[str] = []
        existing = self.state.copy_of(page_id)
        if existing:
            try:
                await self.graph.page(existing)
                return existing, page_id, [f"Editing the existing working copy (id: {existing})."]
            except GraphError as exc:
                if exc.status != 404:
                    raise
                self.state.forget(existing)
                notes.append("The previous working copy was deleted, so a new one was made.")

        job = self.start_copy(page_id, suffix=COPY_SUFFIX, make_working=True)
        await self.wait(job, AUTO_COPY_WAIT)
        if job.status == "running":
            raise EditError(
                "Original pages are never edited, so a working copy is being made first. "
                + job.summary() + " When it's done, repeat this edit; it will go to the copy."
            )
        if job.status == "failed":
            raise EditError(job.summary())
        copy_id = self.state.copy_of(page_id) or job.page_id
        notes.append("First edit of this page: " + job.message + " All edits go to the copy.")
        return copy_id, page_id, notes

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

    def start_restore(self, page_id: str) -> CopyJob:
        """Rebuild the working page from its snapshot as a new page, in the background."""
        working = page_id if self.state.is_working_page(page_id) else self.state.copy_of(page_id)
        if not working:
            raise EditError("This page hasn't been edited by Claude, so there's no snapshot to restore.")
        html = self.state.load_snapshot(working)
        if html is None:
            raise EditError("No snapshot is saved for this page.")
        running = self.running_job_for(working)
        if running:
            return running
        job = CopyJob(id=f"restore-{next(self._job_numbers)}", kind="restore", source_id=working)

        async def work(job: CopyJob) -> None:
            meta = await self.graph.page(working)
            section_id = (meta.get("parentSection") or {}).get("id")
            if not section_id:
                raise EditError("Couldn't find the section this page is in.")
            job.title = meta.get("title") or "Untitled"
            job.method = "rebuilt"
            await self._rebuild(job, section_id, html, job.title)
            original = self.state.original_of(working)
            self.state.forget(working)
            if original:
                self.state.set_copy(original, job.page_id)
            else:
                self.state.add_created(job.page_id)
            self.state.save_snapshot(job.page_id, html)
            job.message = (
                f"Restored '{job.title}' as it was before the last edit, as a new page (id: "
                f"{job.page_id}) in the same section"
                + (f"; {job.failed} image(s)/file(s) couldn't be copied" if job.failed else "")
                + ". Further edits go to the restored page. The page with the bad edit "
                f"(id: {working}) is still there; delete it in OneNote when you've checked the restore."
            )

        return self._start(job, work)
