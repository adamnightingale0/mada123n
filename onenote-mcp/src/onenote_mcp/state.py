"""Local edit state: which pages are working copies, plus one snapshot per page."""

from __future__ import annotations

import hashlib
import json
import logging
from pathlib import Path

from .files import restrict_permissions, write_private

log = logging.getLogger(__name__)


class EditState:
    """Persists original->working-copy links and pages Claude created.

    Stored in ``edit_state.json`` in the app directory. Snapshots live in
    ``snapshots/``, one file per page, overwritten before each edit.
    """

    def __init__(self, base: Path):
        self.path = base / "edit_state.json"
        self.snapshot_dir = base / "snapshots"
        self.copies: dict[str, str] = {}
        self.created: set[str] = set()
        if self.path.exists():
            restrict_permissions(self.path)
            try:
                data = json.loads(self.path.read_text(encoding="utf-8"))
                self.copies = dict(data.get("copies", {}))
                self.created = set(data.get("created", []))
            except (OSError, ValueError) as exc:
                log.warning("Ignoring unreadable edit state %s: %s", self.path, exc)

    def _save(self) -> None:
        data = {"copies": self.copies, "created": sorted(self.created)}
        write_private(self.path, json.dumps(data, indent=2))

    # -- working pages -----------------------------------------------------

    def is_working_page(self, page_id: str) -> bool:
        return page_id in self.created or page_id in self.copies.values()

    def copy_of(self, original_id: str) -> str | None:
        return self.copies.get(original_id)

    def original_of(self, copy_id: str) -> str | None:
        return next((o for o, c in self.copies.items() if c == copy_id), None)

    def set_copy(self, original_id: str, copy_id: str) -> None:
        self.copies[original_id] = copy_id
        self._save()

    def add_created(self, page_id: str) -> None:
        self.created.add(page_id)
        self._save()

    def forget(self, page_id: str) -> None:
        """Drop every record of a page (deleted, or replaced by a restore)."""
        self.copies = {o: c for o, c in self.copies.items() if page_id not in (o, c)}
        self.created.discard(page_id)
        self._save()
        self.snapshot_path(page_id).unlink(missing_ok=True)

    # -- snapshots ---------------------------------------------------------

    def snapshot_path(self, page_id: str) -> Path:
        digest = hashlib.sha256(page_id.encode("utf-8")).hexdigest()[:24]
        return self.snapshot_dir / f"{digest}.html"

    def save_snapshot(self, page_id: str, html: str) -> None:
        write_private(self.snapshot_path(page_id), html)

    def load_snapshot(self, page_id: str) -> str | None:
        path = self.snapshot_path(page_id)
        return path.read_text(encoding="utf-8") if path.exists() else None
