"""Owner-only file helpers for tokens, edit state and page snapshots."""

from __future__ import annotations

import logging
import os
import sys
import tempfile
from pathlib import Path

log = logging.getLogger(__name__)


def restrict_permissions(path: Path) -> None:
    if sys.platform == "win32":
        return
    try:
        os.chmod(path, 0o600)
    except OSError as exc:
        log.warning("Could not restrict permissions on %s: %s", path, exc)


def write_private(path: Path, data: str) -> None:
    """Atomically write ``data`` to ``path`` readable only by the current user.

    On Windows, POSIX modes are mostly ignored; the default app directory
    under %LOCALAPPDATA% is already private to the user profile.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    if sys.platform != "win32":
        os.chmod(path.parent, 0o700)
    fd, tmp = tempfile.mkstemp(dir=path.parent, prefix=f".{path.name}.", suffix=".tmp")
    try:
        if sys.platform != "win32":
            os.fchmod(fd, 0o600)
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            fh.write(data)
        os.replace(tmp, path)
    except BaseException:
        try:
            os.unlink(tmp)
        except FileNotFoundError:
            pass
        raise
