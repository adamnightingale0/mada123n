"""Copy jobs, copy-on-first-edit, target resolution, snapshots and restore."""

from __future__ import annotations

import asyncio
import logging
import re
import time
import uuid
from dataclasses import dataclass, field, fields

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
AUTO_COPY_WAIT = 8.0
DOWNLOAD_CONCURRENCY = 6
BATCH_MAX_ITEMS = 8
BATCH_MAX_BYTES = 20 * 1024 * 1024
PLACEHOLDER_RE = re.compile(r"^\[Copying (image|file) (\d+) of (\d+)(?:…|\.\.\.)\]$")


class EditError(Exception):
    """An edit can't go ahead. The message is for the user/model."""


def _transient(exc: GraphError) -> bool:
    """Connection loss, timeouts, throttling and server errors: worth resuming later."""
    return exc.status is None or exc.status == 429 or exc.status >= 500


def _warning(job: "CopyJob") -> str:
    if not job.problems:
        return ""
    return (
        "⚠️ The copy finished, but a check found problems: " + "; ".join(job.problems)
        + ". Look at the copy in OneNote before relying on it."
    )


INTERRUPTED_HINT = (
    "The copy stopped partway ({detail}). The page keeps its remaining placeholders; call "
    "copy_page again for the same original (or get_copy_status for this job) to finish it."
)


@dataclass
class Prepared:
    page_id: str  # the page the edit is applied to
    target: str | None
    notes: list[str]
    html: str = ""  # the page before the edit, with element IDs


def _is_graph_url(url: str | None) -> bool:
    return bool(url) and url.startswith(("https://graph.microsoft.com/", GRAPH_BASE))


@dataclass
class _Resource:
    kind: str  # "img" or "object"
    url: str
    attrs: dict
    number: int
    mime: str = ""  # type OneNote recorded for it (data-src-type), if any


_MAGIC = [
    (b"\x89PNG\r\n\x1a\n", "image/png"),
    (b"\xff\xd8\xff", "image/jpeg"),
    (b"GIF87a", "image/gif"),
    (b"GIF89a", "image/gif"),
    (b"BM", "image/bmp"),
    (b"II*\x00", "image/tiff"),
    (b"MM\x00*", "image/tiff"),
    (b"%PDF", "application/pdf"),
]


def sniff(data: bytes) -> str | None:
    """Content type from a file's first bytes; None if unrecognised."""
    for magic, mime in _MAGIC:
        if data.startswith(magic):
            return mime
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return "image/webp"
    return None


def placeholder_text(kind: str, number: int, total: int) -> str:
    return f"[Copying {'image' if kind == 'img' else 'file'} {number} of {total}…]"


def collect_resources(html: str) -> tuple[BeautifulSoup, list[_Resource]]:
    """Parse source page HTML, swapping each image/file for a numbered placeholder.

    The numbering is deterministic, so the same source always yields the same
    list; that is what lets an unfinished copy be resumed later.
    """
    soup = BeautifulSoup(html, "html.parser")
    for el in soup.find_all(True):  # IDs belong to the source page
        el.attrs.pop("id", None)
        el.attrs.pop("data-id", None)
    for meta in soup.find_all("meta"):  # e.g. the source's created date
        if meta.get("name"):
            meta.decompose()
    resources: list[_Resource] = []
    placeholders = []
    for el in soup.find_all(["img", "object"]):
        if el.name == "img":
            url, mime = el.get("src"), el.get("data-src-type") or ""
            if not _is_graph_url(url):
                url, mime = el.get("data-fullres-src"), el.get("data-fullres-src-type") or ""
        else:
            url, mime = el.get("data"), el.get("type") or ""
        if not _is_graph_url(url):
            continue
        keep = ("alt", "width", "height", "style") if el.name == "img" else (
            "data-attachment", "type", "style")
        res = _Resource(el.name, url, {k: el[k] for k in keep if el.get(k)}, len(resources) + 1, mime)
        resources.append(res)
        ph = soup.new_tag("p")
        el.replace_with(ph)
        placeholders.append((ph, res))
    for ph, res in placeholders:
        ph.string = placeholder_text(res.kind, res.number, len(resources))
    return soup, resources


def find_placeholders(html: str) -> dict[int, tuple[str, int]]:
    """Unfilled placeholders on a page fetched with includeIDs: {number: (id, total)}."""
    soup = BeautifulSoup(html, "html.parser")
    found = {}
    for p in soup.find_all("p"):
        m = PLACEHOLDER_RE.match(p.get_text(strip=True))
        if m and p.get("id"):
            found[int(m.group(2))] = (p["id"], int(m.group(3)))
    return found


@dataclass
class CopyJob:
    """A page copy (or restore) running in the background.

    It keeps going if the tool call that started it is cancelled, and its
    record is saved, so an interrupted job can be found and resumed.
    """

    id: str
    kind: str  # "copy" or "restore"
    source_id: str  # page copied from (for a restore: the page being undone)
    html_source: str = ""  # "page:<id>" or "snapshot:<page id>": where images come from
    make_working: bool = False
    title: str = ""
    status: str = "running"  # running / done / failed / interrupted
    stage: str = "starting"
    method: str | None = None  # "exact" or "rebuilt"
    page_id: str | None = None
    total: int = 0
    done: int = 0
    failed: int = 0
    note: str = ""
    problems: list = field(default_factory=list)
    message: str = ""
    error: str = ""
    started_at: float = field(default_factory=time.time)
    finished_at: float | None = None
    task: asyncio.Task | None = field(default=None, repr=False, compare=False)

    def record(self) -> dict:
        return {f.name: getattr(self, f.name) for f in fields(self) if f.name != "task"}

    @classmethod
    def from_record(cls, rec: dict) -> "CopyJob":
        names = {f.name for f in fields(cls)} - {"task"}
        return cls(**{k: v for k, v in rec.items() if k in names})

    def progress(self) -> str:
        return f"{self.done}/{self.total}" if self.total else ""

    def summary(self) -> str:
        elapsed = (self.finished_at or time.time()) - self.started_at
        name = f"{self.kind.capitalize()} job {self.id}"
        if self.status == "done":
            return self.message
        if self.status == "failed":
            out = f"{name} failed after {elapsed:.0f}s: {self.error}"
            if self.page_id:
                out += (
                    f" The new page exists (id: {self.page_id}) with {self.done - self.failed} of "
                    f"{self.total} images/files. Calling copy_page for the same original again "
                    "finishes it."
                )
            return out
        if self.status == "interrupted":
            return (
                f"{name} was interrupted (the server restarted) at {self.progress() or 'the start'}."
                + (f" Its page (id: {self.page_id}) is unfinished; calling copy_page or "
                   "get_copy_status again resumes it." if self.page_id else
                   " No page was created; start copy_page again.")
            )
        out = f"{name} is still running ({elapsed:.0f}s so far): {self.stage}."
        if self.page_id:
            out += f" The new page already exists (id: {self.page_id}) and is filling in."
        return out + (
            " It continues in the background even if this chat is interrupted. Call "
            f"get_copy_status with job_id {self.id!r} to follow it; don't start it again."
        )

    def line(self) -> str:
        bits = [self.id, self.kind, self.status]
        if self.progress():
            bits.append(self.progress())
        if self.title:
            bits.append(repr(self.title))
        if self.page_id:
            bits.append(f"page {self.page_id}")
        return " · ".join(bits)


class Editor:
    def __init__(self, graph: GraphClient, state: EditState):
        self.graph = graph
        self.state = state
        self.jobs: dict[str, CopyJob] = {}
        for rec in state.jobs.values():
            job = CopyJob.from_record(rec)
            if job.status == "running":  # its process is gone
                job.status = "interrupted"
                job.stage = "interrupted"
            self.jobs[job.id] = job

    # -- jobs --------------------------------------------------------------

    def _persist(self, job: CopyJob) -> None:
        self.state.save_job(job.record())

    def _start(self, job: CopyJob, work) -> CopyJob:
        self.jobs[job.id] = job
        self._persist(job)

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
                job.finished_at = time.time()
                self._persist(job)
                log.info("%s job %s %s in %.1fs: %s", job.kind, job.id, job.status,
                         job.finished_at - job.started_at, job.message or job.error)

        job.task = asyncio.get_running_loop().create_task(runner())
        return job

    def running_job_for(self, page_id: str) -> CopyJob | None:
        """A running job copying from, or filling in, this page."""
        return next(
            (j for j in self.jobs.values()
             if j.status == "running" and page_id in (j.source_id, j.page_id)), None
        )

    def latest_job_for_page(self, page_id: str) -> CopyJob | None:
        jobs = [j for j in self.jobs.values() if j.page_id == page_id]
        return max(jobs, key=lambda j: j.started_at) if jobs else None

    def recent_jobs(self, limit: int = 10) -> list[CopyJob]:
        return sorted(self.jobs.values(), key=lambda j: j.started_at, reverse=True)[:limit]

    async def wait(self, job: CopyJob, seconds: float) -> CopyJob:
        """Wait up to ``seconds`` for a job; never cancels it."""
        if job.task and not job.task.done() and seconds > 0:
            await asyncio.wait({job.task}, timeout=seconds)
        return job

    @staticmethod
    def _new_id(kind: str) -> str:
        return f"{kind}-{uuid.uuid4().hex[:6]}"

    # -- copying -----------------------------------------------------------

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
        job = CopyJob(id=self._new_id("copy"), kind="copy", source_id=source_id,
                      html_source=f"page:{source_id}", make_working=make_working)

        async def work(job: CopyJob) -> None:
            await self._copy(job, section_id, new_title, suffix)

        return self._start(job, work)

    async def _copy(self, job: CopyJob, section_id: str | None, new_title: str | None, suffix: str) -> None:
        meta = await self.graph.page(job.source_id)
        section_id = section_id or (meta.get("parentSection") or {}).get("id")
        if not section_id:
            raise EditError("Couldn't find the section this page is in.")
        job.title = new_title or (meta.get("title") or "Untitled") + suffix

        exact_ok = False
        if self.state.notes.get("exact_copy_unsupported"):
            job.note = "OneNote's exact copy isn't available for this account, so the page was rebuilt"
        else:
            job.stage = "asking OneNote for an exact copy"

            def waiting(seconds: float) -> None:
                job.stage = f"waiting for OneNote's exact copy ({seconds:.0f}s)"

            try:
                job.page_id = await self.graph.copy_page(
                    job.source_id, section_id, EXACT_COPY_TIMEOUT, on_wait=waiting
                )
                exact_ok = True
            except GraphError as exc:
                if exc.status not in COPY_FALLBACK_STATUSES:
                    raise
                log.warning("Exact copy unavailable (%s); rebuilding the page instead", exc)
                if exc.status == 501:  # "feature not implemented": consumer notebooks
                    self.state.set_note("exact_copy_unsupported", True)
                    job.note = "OneNote's exact copy isn't available for this account, so the page was rebuilt"
                else:
                    job.note = f"OneNote's exact copy didn't work ({exc}), so the page was rebuilt"
                if exc.status == 408:
                    job.note += (
                        ". If a second copy with the original's title appears later, OneNote finished "
                        "its exact copy after all; delete whichever you don't want"
                    )

        if exact_ok:
            job.method = "exact"
            self._persist(job)
            role = self._register_copy(job)
            await self._set_title(job, meta)
        else:
            job.method = "rebuilt"
            job.note += "; handwriting/ink and drawings are not included."
            job.stage = "reading the page"
            html = await self.graph.page_content(job.source_id)
            soup, resources = collect_resources(html)
            if soup.title is not None:
                soup.title.string = job.title
            job.total = len(resources)
            job.stage = "creating the new page"
            page = await self.graph.create_page(section_id, str(soup), [])
            job.page_id = page["id"]
            self._persist(job)
            role = self._register_copy(job)  # registered now, so a retry can't duplicate it
            await self._set_title(job, meta)
            if resources:
                await self._fill_placeholders(job, resources)

        what = (
            "an exact copy (slide printouts, images, attachments and ink included)"
            if job.method == "exact"
            else f"a rebuilt copy with {job.done - job.failed} of {job.total} images/files"
        )
        job.message = " ".join(x for x in (
            _warning(job),
            f"Made {what}: '{job.title}' (id: {job.page_id}).",
            f"{job.failed} image(s)/file(s) couldn't be copied and are marked on the page."
            if job.failed else "",
            job.note,
            role,
        ) if x)

    async def _set_title(self, job: CopyJob, meta: dict) -> None:
        job.stage = "setting the title"
        try:
            await self.graph.update_page(
                job.page_id, [{"target": "title", "action": "replace", "content": job.title}], []
            )
        except GraphError as exc:
            log.warning("Could not set the copy's title: %s", exc)
            if job.method == "exact":
                job.title = meta.get("title") or "Untitled"
                job.note = (job.note + " " if job.note else "") + (
                    f"It couldn't be renamed, so it has the original's title ({exc}).")

    def _register_copy(self, job: CopyJob) -> str:
        source = job.source_id
        if job.make_working and not self.state.is_working_page(source) and not self.state.copy_of(source):
            self.state.set_copy(source, job.page_id)
            return (
                "It is now this page's working copy: edits made with either ID go to the copy, "
                "and the original stays unchanged."
            )
        if self.state.copy_of(source) == job.page_id:
            return "It is this page's working copy."
        self.state.add_created(job.page_id)
        return "It is a separate page; edit it by its ID."

    # -- filling placeholders (also used to resume) --------------------------

    async def _fill_placeholders(self, job: CopyJob, resources: list[_Resource]) -> None:
        """Replace the page's remaining placeholders with the source's images/files.

        Placeholders are found by their text and replaced via the IDs OneNote
        generated for them (paragraph replace doesn't accept data-id targets).
        Downloads run a few at a time; uploads go in batches.
        """
        job.stage = "finding placeholders"
        targets = find_placeholders(await self.graph.page_content(job.page_id, include_ids=True))
        job.total = len(resources)
        if targets and any(total != len(resources) for _, total in targets.values()):
            raise EditError(
                "The original page has changed since this copy was started, so the remaining "
                "images can't be matched up. Delete the unfinished copy and copy the page again."
            )
        todo = [r for r in resources if r.number in targets]
        job.done = job.total - len(todo)
        self._persist(job)

        sem = asyncio.Semaphore(DOWNLOAD_CONCURRENCY)

        async def fetch(res: _Resource):
            async with sem:
                try:
                    data, header = await self.graph.get_bytes(res.url)
                except GraphError as exc:
                    if _transient(exc):
                        raise
                    log.warning("Couldn't download %s %d: %s", res.kind, res.number, exc)
                    return res, None, str(exc)
                header = header.split(";")[0].strip()
                detected = sniff(data)
                log.info("Downloaded %s %d: %d bytes, header %s, recorded %s, detected %s",
                         res.kind, res.number, len(data), header, res.mime or "-", detected or "-")
                if res.kind == "img":
                    if not (detected or "").startswith("image/"):
                        log.warning("%s %d isn't image data; first bytes: %r",
                                    res.kind, res.number, data[:120])
                        return res, None, "not image data"
                    return res, data, detected
                ctype = res.mime or detected or header or "application/octet-stream"
                return res, data, ctype

        for i in range(0, len(todo), BATCH_MAX_ITEMS):
            chunk = todo[i:i + BATCH_MAX_ITEMS]
            job.stage = f"copying images and files ({job.done}/{job.total})"
            try:
                fetched = await asyncio.gather(*(fetch(r) for r in chunk))
            except GraphError as exc:
                raise EditError(INTERRUPTED_HINT.format(detail=exc)) from exc
            batch, size = [], 0
            for res, data, ctype in fetched:
                target = targets[res.number][0]
                if data is None:
                    await self._fill(job, [(res, target, None, "")])
                    continue
                if batch and size + len(data) > BATCH_MAX_BYTES:
                    await self._fill(job, batch)
                    batch, size = [], 0
                batch.append((res, target, data, ctype))
                size += len(data)
            if batch:
                await self._fill(job, batch)
            self._persist(job)
        job.stage = "checking the copy"
        html = await self.graph.page_content(job.page_id, include_ids=True)
        left = find_placeholders(html)
        if left:
            raise EditError(INTERRUPTED_HINT.format(
                detail=f"{len(left)} image(s)/file(s) still to copy"))
        job.problems = await self._verify(html, resources)
        if job.problems:
            log.warning("Copy check found problems: %s", "; ".join(job.problems))
        job.stage = "finished"

    async def _verify(self, html: str, resources: list[_Resource]) -> list[str]:
        """Download a few images back from the copy and check they are real images."""
        soup = BeautifulSoup(html, "html.parser")
        imgs = [i for i in soup.find_all("img") if _is_graph_url(i.get("src"))]
        objects = [o for o in soup.find_all("object") if _is_graph_url(o.get("data"))]
        problems = []
        want_imgs = sum(1 for r in resources if r.kind == "img")
        want_objects = len(resources) - want_imgs
        if want_objects and not objects:
            problems.append(f"the copy has no attached files, but the original has {want_objects}")
        sample = sorted({0, len(imgs) // 2, len(imgs) - 1}) if imgs else []
        bad = 0
        for i in sample:
            try:
                data, header = await self.graph.get_bytes(imgs[i]["src"])
            except GraphError as exc:
                if _transient(exc):
                    raise
                bad += 1
                log.warning("Check: image %d on the copy can't be downloaded: %s", i + 1, exc)
                continue
            detected = sniff(data)
            log.info("Check: image %d on the copy: %d bytes, header %s, detected %s",
                     i + 1, len(data), header, detected or "-")
            if not (detected or "").startswith("image/"):
                bad += 1
        if bad:
            problems.append(f"{bad} of {len(sample)} sampled images on the copy aren't readable images")
        if want_imgs and len(imgs) < want_imgs:
            problems.append(f"the copy shows {len(imgs)} images, the original {want_imgs}")
        return problems

    async def _fill(self, job: CopyJob, items: list[tuple[_Resource, str, bytes | None, str]]) -> None:
        commands, parts = [], []
        for res, target, data, ctype in items:
            if data is None:
                commands.append(self._failed_command(res, target))
                job.failed += 1
                continue
            name = f"{'img' if res.kind == 'img' else 'file'}{res.number}"
            parts.append(Part(name, data, ctype))
            tag = BeautifulSoup("", "html.parser").new_tag(res.kind, attrs=dict(res.attrs))
            tag["src" if res.kind == "img" else "data"] = f"name:{name}"
            commands.append({"target": target, "action": "replace", "content": str(tag)})
        try:
            await self.graph.update_page(job.page_id, commands, parts)
        except GraphError as exc:
            if exc.status == 404:
                raise EditError("The copy was deleted while it was being filled in.") from exc
            if _transient(exc):
                raise EditError(INTERRUPTED_HINT.format(detail=exc)) from exc
            if len(items) > 1:
                log.warning("Batch upload failed (%s); retrying items one by one", exc)
                for item in items:
                    await self._fill(job, [item])
                return
            res, target = items[0][0], items[0][1]
            log.warning("Upload of %s %d failed: %s", res.kind, res.number, exc)
            if items[0][2] is not None:
                job.failed += 1
            try:
                await self.graph.update_page(job.page_id, [self._failed_command(res, target)], [])
            except GraphError:
                pass
        job.done += len(items)
        job.stage = f"copying images and files ({job.done}/{job.total})"

    @staticmethod
    def _failed_command(res: _Resource, target: str) -> dict:
        kind = "Image" if res.kind == "img" else "Attached file"
        label = res.attrs.get("data-attachment", "")
        return {"target": target, "action": "replace",
                "content": f"<p>[{kind} {res.number} could not be copied{': ' + label if label else ''}]</p>"}

    async def ensure_filled(self, page_id: str, html: str | None = None) -> CopyJob | None:
        """If a copy is unfinished, return the job filling it (resuming it if needed)."""
        running = self.running_job_for(page_id)
        if running:
            return running
        if html is None:
            html = await self.graph.page_content(page_id, include_ids=True)
        if not find_placeholders(html):
            return None
        previous = self.latest_job_for_page(page_id)
        html_source = previous.html_source if previous and previous.html_source else ""
        if not html_source:
            original = self.state.original_of(page_id)
            if not original:
                raise EditError("This page has unfinished copy placeholders, but its original is unknown.")
            html_source = f"page:{original}"
        job = CopyJob(
            id=self._new_id("copy"), kind=previous.kind if previous else "copy",
            source_id=previous.source_id if previous else html_source.split(":", 1)[1],
            html_source=html_source, page_id=page_id, method="rebuilt",
            title=previous.title if previous else "",
        )

        async def work(job: CopyJob) -> None:
            if not job.title:
                job.title = (await self.graph.page(page_id)).get("title") or ""
            kind, ref = job.html_source.split(":", 1)
            src = await self.graph.page_content(ref) if kind == "page" else self.state.load_snapshot(ref)
            if src is None:
                raise EditError("The snapshot this page was restored from is gone; it can't be finished.")
            _, resources = collect_resources(src)
            await self._fill_placeholders(job, resources)
            job.message = _warning(job) + (" " if job.problems else "") + (
                f"Finished copying the remaining images/files into '{job.title or 'the copy'}' "
                f"(id: {page_id}): {job.done - job.failed} of {job.total} in place"
                + (f", {job.failed} couldn't be copied and are marked on the page" if job.failed else "")
                + "."
            )

        log.info("Resuming unfinished copy %s", page_id)
        return self._start(job, work)

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
        filling = await self.ensure_filled(working, html)
        if filling:
            await self.wait(filling, AUTO_COPY_WAIT)
            if filling.status != "done":
                raise EditError(
                    "This page's images and files are still being copied in, so it can't be edited "
                    "yet. " + filling.summary()
                )
            html = await self.graph.page_content(working, include_ids=True)
        self.state.save_snapshot(working, html)

        if target in (None, "title", "body"):
            return Prepared(working, target, notes, html)
        structure = page_structure(html)
        if any(e.target == target for e in structure):
            return Prepared(working, target, notes, html)
        if original:
            # The ID came from the original page; find the same element in the copy.
            orig = page_structure(await self.graph.page_content(original, include_ids=True))
            idx = next((i for i, e in enumerate(orig) if e.target == target), None)
            if idx is not None and idx < len(structure):
                cand = structure[idx]
                if cand.tag == orig[idx].tag and cand.preview == orig[idx].preview:
                    return Prepared(working, cand.target, notes, html)
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
        job = CopyJob(id=self._new_id("restore"), kind="restore", source_id=working, method="rebuilt")

        async def work(job: CopyJob) -> None:
            meta = await self.graph.page(working)
            section_id = (meta.get("parentSection") or {}).get("id")
            if not section_id:
                raise EditError("Couldn't find the section this page is in.")
            job.title = meta.get("title") or "Untitled"
            soup, resources = collect_resources(html)
            if soup.title is not None:
                soup.title.string = job.title
            job.total = len(resources)
            job.stage = "creating the restored page"
            page = await self.graph.create_page(section_id, str(soup), [])
            job.page_id = page["id"]
            # Re-point state now, so a retry or a resume finds the new page.
            self.state.save_snapshot(job.page_id, html)
            job.html_source = f"snapshot:{job.page_id}"
            original = self.state.original_of(working)
            self.state.forget(working)
            if original:
                self.state.set_copy(original, job.page_id)
            else:
                self.state.add_created(job.page_id)
            self._persist(job)
            await self._set_title(job, meta)
            if resources:
                await self._fill_placeholders(job, resources)
            job.message = _warning(job) + (" " if job.problems else "") + (
                f"Restored '{job.title}' as it was before the last edit, as a new page (id: "
                f"{job.page_id}) in the same section"
                + (f"; {job.failed} image(s)/file(s) couldn't be copied" if job.failed else "")
                + ". Further edits go to the restored page. The page with the bad edit "
                f"(id: {working}) is still there; delete it in OneNote when you've checked the restore."
            )

        return self._start(job, work)
