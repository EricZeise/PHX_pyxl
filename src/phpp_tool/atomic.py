"""Atomic file output: write to a temporary file beside the destination, then replace the destination in one step.

A crash, a full disk or an interrupted run while writing leaves the previous file exactly as it was (and no
half-written file where a new one was expected); the temporary file is removed on any failure. The destination
keeps its permissions when it already exists. Used for every file this package produces: the JSON building
record, the written PHPP workbook, and the comparison reports.
"""

from __future__ import annotations

import contextlib
import os
import shutil
import stat
import tempfile
from pathlib import Path

_NEW_FILE_MODE = (
    0o644  # mkstemp creates 0600; a new output file gets the usual readable mode
)


def _temp_beside(path: Path) -> tuple[int, Path]:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(dir=path.parent, prefix=f".{path.name}.", suffix=".tmp")
    return fd, Path(name)


def _mode_of(path: Path) -> int:
    try:
        return stat.S_IMODE(path.stat().st_mode)
    except FileNotFoundError:
        return _NEW_FILE_MODE


def atomic_write_bytes(path: str | Path, data: bytes) -> None:
    """Write `data` to `path` atomically (the parent folder is created if missing)."""
    path = Path(path)
    fd, tmp = _temp_beside(path)
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(data)
            f.flush()
            os.fsync(f.fileno())
        os.chmod(tmp, _mode_of(path))
        os.replace(tmp, path)
    except BaseException:
        with contextlib.suppress(OSError):
            tmp.unlink()
        raise


def atomic_write_text(path: str | Path, text: str, encoding: str = "utf-8") -> None:
    """Write `text` to `path` atomically."""
    atomic_write_bytes(path, text.encode(encoding))


def atomic_copy(source: str | Path, destination: str | Path) -> None:
    """Copy `source` to `destination` (contents and timestamps, like `shutil.copy2`) atomically."""
    destination = Path(destination)
    fd, tmp = _temp_beside(destination)
    os.close(fd)
    try:
        shutil.copy2(source, tmp)
        os.chmod(
            tmp,
            _mode_of(destination)
            if destination.exists()
            else stat.S_IMODE(Path(source).stat().st_mode),
        )
        os.replace(tmp, destination)
    except BaseException:
        with contextlib.suppress(OSError):
            tmp.unlink()
        raise
