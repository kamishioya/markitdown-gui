from __future__ import annotations

import os
import re
import shutil
import tempfile
import time
from pathlib import Path

COPILOT_TEMP_PREFIX = "markitdown-copilot-"
XLSX_PDF_SCRIPT_TEMP_PREFIX = "markitdown-xlsx-pdf-script-"
XLSX_PDF_TEMP_PREFIX = "markitdown-gui-xlsx-pdf-"
MARKITDOWN_TEMP_PREFIXES = (
    XLSX_PDF_SCRIPT_TEMP_PREFIX,
    XLSX_PDF_TEMP_PREFIX,
    COPILOT_TEMP_PREFIX,
)
LEGACY_TEMP_DIR_STALE_AFTER_SECONDS = 24 * 60 * 60

_PID_SEGMENT_RE = re.compile(r"^pid(?P<pid>\d+)-")


def build_temp_dir_prefix(base_prefix: str) -> str:
    return f"{base_prefix}pid{os.getpid()}-"


def cleanup_markitdown_temp_dirs(
    *,
    temp_root: str | Path | None = None,
    legacy_stale_after_seconds: int = LEGACY_TEMP_DIR_STALE_AFTER_SECONDS,
    current_time: float | None = None,
) -> list[Path]:
    root = Path(temp_root) if temp_root is not None else Path(tempfile.gettempdir())
    now = time.time() if current_time is None else current_time
    removed_paths: list[Path] = []

    if not root.exists():
        return removed_paths

    try:
        entries = list(root.iterdir())
    except OSError:
        return removed_paths

    for entry in entries:
        if not entry.is_dir():
            continue

        prefix = _matching_temp_prefix(entry.name)
        if prefix is None:
            continue

        if not _should_remove_temp_dir(
            entry,
            prefix,
            now,
            legacy_stale_after_seconds,
        ):
            continue

        try:
            shutil.rmtree(entry)
        except OSError:
            continue

        removed_paths.append(entry)

    return removed_paths


def _matching_temp_prefix(directory_name: str) -> str | None:
    for prefix in MARKITDOWN_TEMP_PREFIXES:
        if directory_name.startswith(prefix):
            return prefix
    return None


def _should_remove_temp_dir(
    directory_path: Path,
    prefix: str,
    current_time: float,
    legacy_stale_after_seconds: int,
) -> bool:
    pid = _extract_pid(directory_path.name, prefix)
    if pid is not None:
        return not _is_process_running(pid)

    try:
        modified_at = directory_path.stat().st_mtime
    except OSError:
        return False

    return current_time - modified_at >= legacy_stale_after_seconds


def _extract_pid(directory_name: str, prefix: str) -> int | None:
    remainder = directory_name[len(prefix) :]
    match = _PID_SEGMENT_RE.match(remainder)
    if match is None:
        return None
    return int(match.group("pid"))


def _is_process_running(pid: int) -> bool:
    if pid <= 0:
        return False

    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    except OSError:
        return False

    return True