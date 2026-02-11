#!/usr/bin/env python3
"""
Read INSTRUCTION.xlsx and print from/to time and 15-minute slots for each entry.

15-minute division matches the calculation sheet:
  "calculation sheet for BD and non compliance of HNPCL for OCT25.xlsx"
  (sheet "New method (2)"): each slot is a row with From and To columns,
  e.g. 2:30–2:45, 2:45–3:00, 3:00–3:15, ... (consecutive 15-min intervals).
  Day is divided as 00:00, 00:15, 00:30, 00:45, 01:00, ... 23:45.
"""

import datetime
import sys
from pathlib import Path


def format_time(val):
    """Format time or datetime for display."""
    if val is None or val == "":
        return ""
    if isinstance(val, datetime.time):
        return val.strftime("%H:%M")
    if isinstance(val, datetime.datetime):
        return val.strftime("%H:%M")
    return str(val)


def format_date(val):
    """Format date or datetime for display (DD-MMM-YYYY, e.g. 05-Nov-2025)."""
    if val is None or val == "":
        return ""
    if isinstance(val, datetime.datetime):
        return val.strftime("%d-%b-%Y")
    if hasattr(val, "strftime"):
        return val.strftime("%d-%b-%Y")
    return str(val)


def time_to_minutes(t):
    """Convert time or datetime to minutes since midnight (0–1439). Returns None if invalid."""
    if t is None:
        return None
    if isinstance(t, datetime.datetime):
        t = t.time()
    if isinstance(t, datetime.time):
        return t.hour * 60 + t.minute
    return None


def floor_to_15(minutes):
    """Floor minutes since midnight to previous 15-min slot (0, 15, 30, 45, ...)."""
    if minutes is None:
        return None
    return (minutes // 15) * 15


def slots_15min(from_time, to_time):
    """
    Generate 15-minute slots between from_time and to_time.
    Matches calculation sheet: each slot is (From, To) with To = From + 15 min.
    Start/end are floored to previous 15-min boundary (e.g. 8:10 → 8:00).
    Returns list of (from_str, to_str) e.g. [("8:00", "8:15"), ("8:15", "8:30"), ...].
    Handles overnight (e.g. 23:00 to 00:00).
    """
    start_min = time_to_minutes(from_time)
    end_min = time_to_minutes(to_time)
    if start_min is None or end_min is None:
        return []

    start_slot = floor_to_15(start_min)
    end_slot = floor_to_15(end_min)

    if start_slot == end_slot:
        from_str = minutes_to_time_str(start_slot)
        to_str = minutes_to_time_str((start_slot + 15) % (24 * 60))
        return [(from_str, to_str)]

    # Overnight: end is next day (e.g. 23:00 → 00:00)
    if start_slot > end_slot:
        end_slot += 24 * 60

    result = []
    m = start_slot
    while m <= end_slot:
        from_m = m % (24 * 60)
        to_m = (m + 15) % (24 * 60)
        result.append((minutes_to_time_str(from_m), minutes_to_time_str(to_m)))
        m += 15
    return result


def minutes_to_time_str(minutes):
    """Convert minutes since midnight (0–1439) to HH:MM string."""
    h = minutes // 60
    m = minutes % 60
    return f"{h:02d}:{m:02d}"


def slots_to_vertical_mini_table(slots, cell_w=6, use_outer_border=False):
    """
    Build a vertical two-column mini table (From | To) for 15-min slots.
    By default uses the main table's cell as the outline (no inner box).
    With use_outer_border=True, draws full inner table borders.
    """
    if not slots:
        return ["(none)"]
    lines = []
    inner_sep = f"{'-' * (cell_w + 2)}+{'-' * (cell_w + 2)}"   # between columns only
    if use_outer_border:
        full_sep = f"+{'-' * (cell_w + 2)}+{'-' * (cell_w + 2)}+"
        lines.append(full_sep)
    header = f" {'From':<{cell_w}} | {'To':<{cell_w}} "
    lines.append(header)
    lines.append(inner_sep)
    for from_str, to_str in slots:
        row = f" {from_str:<{cell_w}} | {to_str:<{cell_w}} "
        lines.append(row)
    if use_outer_border:
        lines.append(full_sep)
    return lines

try:
    import openpyxl
except ImportError:
    print("Install openpyxl: pip install openpyxl", file=sys.stderr)
    sys.exit(1)


def find_time_columns(sheet, max_header_rows=4):
    """
    Find 'from time' and 'to time' column indices (0-based).
    Checks first max_header_rows; supports multi-row headers and merged cells.
    Returns (from_idx, to_idx, data_start_row) or (None, None, 1).
    """
    time_cols = []
    for row_num in range(1, max_header_rows + 1):
        row_cells = list(sheet.iter_rows(min_row=row_num, max_row=row_num))[0]
        for idx, cell in enumerate(row_cells):
            if cell.value is None:
                continue
            val = str(cell.value).strip().lower()
            if "time" in val and idx not in [c for c, _ in time_cols]:
                time_cols.append((idx, val))
    # Data starts after the last header row we scanned
    data_start = max_header_rows + 1
    if len(time_cols) >= 2:
        return time_cols[0][0], time_cols[1][0], data_start
    if len(time_cols) == 1:
        return time_cols[0][0], time_cols[0][0], data_start
    return None, None, 1


def find_date_column(sheet, max_header_rows=4):
    """Find first 'date' column index (0-based) in header rows. Returns None if not found."""
    for row_num in range(1, max_header_rows + 1):
        row_cells = list(sheet.iter_rows(min_row=row_num, max_row=row_num))[0]
        for idx, cell in enumerate(row_cells):
            if cell.value is None:
                continue
            val = str(cell.value).strip().lower()
            if "date" in val:
                return idx
    return None


def main():
    base = Path(__file__).resolve().parent
    for name in ("INSTRUCTION.xlsx", "instructions.xlsx", "Instruction.xlsx"):
        path = base / name
        if path.exists():
            break
    else:
        print("No INSTRUCTION.xlsx or instructions.xlsx found.", file=sys.stderr)
        sys.exit(1)

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet = wb.active

    from_idx, to_idx, data_start_row = find_time_columns(sheet)
    if from_idx is None:
        print("Could not find 'from time' / 'to time' columns in first 5 rows.", file=sys.stderr)
        wb.close()
        sys.exit(1)

    date_idx = find_date_column(sheet, max_header_rows=4)

    # Build data: one entry per instruction row (row_num, date, from, to, slots)
    # slots = list of (from_str, to_str) for each 15-min interval
    rows_data = []
    slot_cell_w = 6
    w_slots_col = 0
    w_date = 10  # YYYY-MM-DD
    for row_num, row in enumerate(
        sheet.iter_rows(min_row=data_start_row, values_only=True), start=data_start_row
    ):
        row = list(row)
        date_val = row[date_idx] if date_idx is not None and date_idx < len(row) else ""
        from_time = row[from_idx] if from_idx < len(row) else ""
        to_time = row[to_idx] if to_idx < len(row) else ""
        slots = slots_15min(from_time, to_time)
        date_str = format_date(date_val)
        fr_str = format_time(from_time)
        to_str = format_time(to_time)
        mini = slots_to_vertical_mini_table(slots, slot_cell_w)
        w_slots_col = max(w_slots_col, max(len(ln) for ln in mini))
        if date_str:
            w_date = max(w_date, len(date_str))
        rows_data.append((row_num, date_str, fr_str, to_str, slots))

    col_row = "Row"
    col_date = "Date"
    col_from = "From"
    col_to = "To"
    col_slots = "Slots"
    w_row = max(len(col_row), 4)
    w_date = max(len(col_date), w_date)
    w_from = max(len(col_from), 6)
    w_to = max(len(col_to), 6)
    w_slots_col = max(len(col_slots), w_slots_col)

    print(f"Reading: {path.name}\n")

    # Table: Row | Date | From | To | Slots (vertical From|To mini table inside each row's Slots cell)
    sep_main = f"+{'-' * (w_row + 2)}+{'-' * (w_date + 2)}+{'-' * (w_from + 2)}+{'-' * (w_to + 2)}+{'-' * (w_slots_col + 2)}+"
    head = f"| {col_row:<{w_row}} | {col_date:<{w_date}} | {col_from:<{w_from}} | {col_to:<{w_to}} | {col_slots:<{w_slots_col}} |"
    print(sep_main)
    print(head)
    print(sep_main)

    for row_num, date_str, fr, to, slots in rows_data:
        mini_lines = slots_to_vertical_mini_table(slots, slot_cell_w)
        for line_idx, mt_line in enumerate(mini_lines):
            slot_cell = mt_line.ljust(w_slots_col) if len(mt_line) <= w_slots_col else mt_line
            if line_idx == 0:
                print(f"| {row_num:<{w_row}} | {date_str:<{w_date}} | {fr:<{w_from}} | {to:<{w_to}} | {slot_cell} |")
            else:
                print(f"| {'':<{w_row}} | {'':<{w_date}} | {'':<{w_from}} | {'':<{w_to}} | {slot_cell} |")
        print(sep_main)

    wb.close()


if __name__ == "__main__":
    main()
