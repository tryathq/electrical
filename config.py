"""App configuration and constants."""

from pathlib import Path

# Paths
APP_DIR = Path(__file__).resolve().parent
REPORTS_DIR = APP_DIR / "reports"
REPORTS_INDEX_FILE = REPORTS_DIR / "reports_index.json"
BACKGROUND_JOB_FILE = APP_DIR / "background_job.json"

# Processing
PROCESSING_BATCH_SIZE = 5
# Write partial_output.json every N slots (job progress still every PROCESSING_BATCH_SIZE)
PARTIAL_OUTPUT_WRITE_INTERVAL = 25

# Table display
TABLE_ROW_PX = 35
TABLE_HEADER_PX = 40
TABLE_MIN_HEIGHT = 300
# Approximate remaining viewport height after Streamlit chrome + header/stats/filters
# Typical screen 900-1080px minus ~350-400px for elements above table
TABLE_VIEWPORT_HEIGHT = 550  # Adjust this based on your screen
PAGE_SIZE_ALL = "ALL"  # Label for "show all rows" option in pagination


def table_height(row_count: int) -> int:
    """Compute table height to fill remaining viewport.
    
    For small tables: size to content
    For large tables: use viewport height (table scrolls internally)
    """
    content_height = TABLE_HEADER_PX + row_count * TABLE_ROW_PX
    # If content fits in less than viewport, use content height
    if content_height < TABLE_VIEWPORT_HEIGHT:
        return max(TABLE_MIN_HEIGHT, content_height)
    # Otherwise fill available viewport space
    return TABLE_VIEWPORT_HEIGHT
