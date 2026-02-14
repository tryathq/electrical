"""Reports persistence: save/load report index and files."""

import json
import os
import shutil
from pathlib import Path

from config import REPORTS_DIR, REPORTS_INDEX_FILE


def ensure_dir() -> None:
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)


def load_index() -> list:
    """Load list of persisted reports from disk (newest first)."""
    if not REPORTS_INDEX_FILE.exists():
        return []
    try:
        with open(REPORTS_INDEX_FILE, "r", encoding="utf-8") as f:
            entries = json.load(f)
        return sorted(entries, key=lambda e: e.get("run_at", ""), reverse=True)
    except Exception:
        return []


def append_entry(entry: dict) -> None:
    """Append one report entry to the index and flush to disk."""
    ensure_dir()
    entries = load_index()
    entries.insert(0, entry)
    with open(REPORTS_INDEX_FILE, "w", encoding="utf-8") as f:
        json.dump(entries, f, indent=2, ensure_ascii=False)
        f.flush()
        os.fsync(f.fileno())


def save_file(src_path: Path, filename: str) -> Path:
    """Copy report file to reports dir; returns destination path."""
    ensure_dir()
    dest = REPORTS_DIR / filename
    shutil.copy2(src_path, dest)
    return dest
