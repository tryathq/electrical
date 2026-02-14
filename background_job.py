"""Background report generation: read/write job state to disk so it survives navigation."""

import json
import os
from pathlib import Path

from config import BACKGROUND_JOB_FILE


def read_job() -> dict | None:
    """Read current background job state from disk. Returns None if no job or file missing."""
    if not BACKGROUND_JOB_FILE.exists():
        return None
    try:
        with open(BACKGROUND_JOB_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def write_job(data: dict) -> None:
    """Write job state to disk (thread-safe enough for progress updates)."""
    BACKGROUND_JOB_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(BACKGROUND_JOB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
        f.flush()
        os.fsync(f.fileno())
