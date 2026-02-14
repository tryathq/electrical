"""App configuration and constants."""

from pathlib import Path

# Paths
APP_DIR = Path(__file__).resolve().parent
REPORTS_DIR = APP_DIR / "reports"
REPORTS_INDEX_FILE = REPORTS_DIR / "reports_index.json"
BACKGROUND_JOB_FILE = APP_DIR / "background_job.json"

# Processing
PROCESSING_BATCH_SIZE = 5

# Table display
TABLE_ROW_PX = 35
TABLE_HEADER_PX = 40
TABLE_MIN_HEIGHT = 200
TABLE_MAX_HEIGHT = 500


def table_height(row_count: int) -> int:
    """Compute table height in pixels from row count."""
    return min(
        max(TABLE_MIN_HEIGHT, TABLE_HEADER_PX + row_count * TABLE_ROW_PX),
        TABLE_MAX_HEIGHT,
    )
