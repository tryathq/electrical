#!/usr/bin/env python3
"""
Find rows in XLSX file where "Name of the station" column matches given station name.

Usage:
  python find_station_rows.py <xlsx_path> <station_name> [options]

Example:
  python find_station_rows.py "input/Back_Down_Instructions.xlsx" HINDUJA
  python find_station_rows.py "input/jan 2026.xlsx" HINDUJA --sheet HNPCL
  python find_station_rows.py "input/instructions.xlsx" HINDUJA --dc-file "input/dc_data.xlsx"
"""

import argparse
import sys
from pathlib import Path
from datetime import datetime, date, time

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side
except ImportError:
    print("Install openpyxl: pip install openpyxl", file=sys.stderr)
    sys.exit(1)


def format_value(val):
    """Format cell value for display, especially dates and times."""
    if val is None:
        return ""
    if isinstance(val, datetime):
        # Check if it's actually a time (1900-01-01 date with time)
        if val.date() == date(1900, 1, 1):
            return val.strftime("%H:%M")
        return val.strftime("%d-%b-%Y")
    if isinstance(val, date):
        return val.strftime("%d-%b-%Y")
    # Check if it's a time object
    if isinstance(val, time):
        return val.strftime("%H:%M")
    # Check if string looks like a time (HH:MM:SS or HH:MM)
    val_str = str(val)
    if ":" in val_str and len(val_str) <= 8:
        # Try to parse and format as time
        try:
            if len(val_str.split(":")) == 3:
                # HH:MM:SS format
                parts = val_str.split(":")
                return f"{parts[0]}:{parts[1]}"
            return val_str
        except:
            pass
    return str(val)


def time_to_minutes(t):
    """Convert time or datetime to minutes since midnight (0–1439). Returns None if invalid."""
    if t is None:
        return None
    if isinstance(t, datetime):
        t = t.time()
    if isinstance(t, time):
        return t.hour * 60 + t.minute
    # Try parsing string
    if isinstance(t, str):
        try:
            parts = t.strip().split(":")
            if len(parts) >= 2:
                h = int(parts[0])
                m = int(parts[1])
                if 0 <= h < 24 and 0 <= m < 60:
                    return h * 60 + m
        except (ValueError, IndexError):
            pass
    return None


def floor_to_15(minutes):
    """Floor minutes since midnight to previous 15-min slot (0, 15, 30, 45, ...)."""
    if minutes is None:
        return None
    return (minutes // 15) * 15


def minutes_to_time_str(minutes):
    """Convert minutes since midnight (0–1439) to HH:MM string."""
    h = minutes // 60
    m = minutes % 60
    return f"{h:02d}:{m:02d}"


def slots_15min(from_time, to_time):
    """
    Generate 15-minute slots between from_time and to_time.
    Start is floored to previous 15-min boundary (e.g. 8:10 → 8:00).
    End time is respected - slots do not extend beyond the given end time.
    Returns list of (from_str, to_str) e.g. [("8:00", "8:15"), ("8:15", "8:30"), ...].
    Handles overnight (e.g. 23:00 to 00:00).
    """
    start_min = time_to_minutes(from_time)
    end_min = time_to_minutes(to_time)
    if start_min is None or end_min is None:
        return []

    start_slot = floor_to_15(start_min)
    
    # Overnight: end is next day (e.g. 23:00 → 00:00)
    if start_slot > end_min:
        end_min += 24 * 60

    result = []
    m = start_slot
    # Only include slots where the slot's START time is < end_min
    # This ensures slots don't extend beyond the end time
    while m < end_min:
        from_m = m % (24 * 60)
        to_m = (m + 15) % (24 * 60)
        result.append((minutes_to_time_str(from_m), minutes_to_time_str(to_m)))
        m += 15
    return result


def parse_time_str(time_str):
    """Parse HH:MM string to minutes since midnight. Returns None if invalid."""
    return time_to_minutes(time_str)


def convert_date_to_sheet_format(date_str):
    """
    Convert date string from format '02-Jan-2026' to sheet format '02.01.2026'.
    Handles various date formats and returns None if parsing fails.
    """
    if not date_str:
        return None
    
    date_str = str(date_str).strip()
    
    # List of formats to try
    formats = [
        "%d-%b-%Y",      # 02-Jan-2026
        "%d-%b-%y",      # 02-Jan-26
        "%d.%m.%Y",      # 02.01.2026 (already in correct format)
        "%d/%m/%Y",      # 02/01/2026
        "%Y-%m-%d",      # 2026-01-02
    ]
    
    for fmt in formats:
        try:
            dt = datetime.strptime(date_str, fmt)
            # Convert to DD.MM.YYYY format
            result = dt.strftime("%d.%m.%Y")
            return result
        except ValueError:
            continue
    
    return None


def normalize_time_str(time_str):
    """
    Normalize time string to HH:MM format for comparison.
    Handles formats like "0:00", "00:00", "1:15", "01:15", etc.
    """
    if not time_str:
        return ""
    time_str = str(time_str).strip()
    # Parse and reformat to ensure consistent HH:MM format
    minutes = time_to_minutes(time_str)
    if minutes is not None:
        return minutes_to_time_str(minutes)
    return time_str


def find_dc_value(dc_wb, sheet_name, from_time_str, to_time_str, debug=False):
    """
    Find DC value from DC workbook sheet for matching time range.
    Searches for row where 'From' and 'To' columns match the given time range.
    Returns the 'Final Revison' column value, or None if not found.
    """
    if dc_wb is None:
        if debug:
            print(f"  [DC Lookup] dc_wb is None", file=sys.stderr)
        return None
    
    if not sheet_name:
        if debug:
            print(f"  [DC Lookup] sheet_name is empty", file=sys.stderr)
        return None
    
    # Find the sheet by name (case-insensitive, flexible matching)
    target_sheet = None
    sheet_name_lower = sheet_name.lower().strip()
    
    # Try exact match first (with date format)
    for name in dc_wb.sheetnames:
        name_clean = name.strip()
        if name_clean.lower() == sheet_name_lower:
            target_sheet = name
            break
    
    # If not found, try partial match (date might be embedded in sheet name)
    if target_sheet is None:
        for name in dc_wb.sheetnames:
            name_clean = name.strip().lower()
            # Check if the date string is contained in the sheet name
            if sheet_name_lower in name_clean:
                target_sheet = name
                break
    
    if target_sheet is None:
        if debug:
            print(f"  [DC Lookup] Sheet '{sheet_name}' not found.", file=sys.stderr)
            print(f"  [DC Lookup] Looking for: '{sheet_name}' (normalized: '{sheet_name_lower}')", file=sys.stderr)
            print(f"  [DC Lookup] Available sheets ({len(dc_wb.sheetnames)}):", file=sys.stderr)
            for i, name in enumerate(dc_wb.sheetnames[:10]):
                print(f"    {i+1}. '{name}'", file=sys.stderr)
            if len(dc_wb.sheetnames) > 10:
                print(f"    ... and {len(dc_wb.sheetnames) - 10} more", file=sys.stderr)
        return None
    
    if debug:
        print(f"  [DC Lookup] ✓ Found sheet: '{target_sheet}' (looking for: '{sheet_name}')", file=sys.stderr)
    
    ws = dc_wb[target_sheet]
    
    # Find column indices for "From", "To", and "Final Revison"
    from_col = None
    to_col = None
    final_revision_col = None
    header_row = None
    
    # First pass: find the header row by looking for a row that has multiple expected headers
    for row_num in range(1, min(11, ws.max_row + 1)):
        found_headers = []
        for col_idx in range(1, min(ws.max_column + 1, 20)):
            cell = ws.cell(row=row_num, column=col_idx)
            if cell.value:
                header_val = str(cell.value).strip().lower()
                if "from" in header_val:
                    found_headers.append(("from", col_idx))
                elif "to" in header_val and "tb" not in header_val and "no" not in header_val:
                    found_headers.append(("to", col_idx))
                elif "final" in header_val and ("revis" in header_val or "revisi" in header_val):
                    found_headers.append(("final", col_idx))
        
        # If we found at least 2 of the expected headers in this row, use it as header row
        if len(found_headers) >= 2:
            header_row = row_num
            for header_type, col_idx in found_headers:
                if header_type == "from" and from_col is None:
                    from_col = col_idx
                elif header_type == "to" and to_col is None:
                    to_col = col_idx
                elif header_type == "final" and final_revision_col is None:
                    final_revision_col = col_idx
            break
    
    # If we didn't find a good header row, do a second pass with less strict matching
    if header_row is None:
        for row_num in range(1, min(11, ws.max_row + 1)):
            for col_idx in range(1, min(ws.max_column + 1, 20)):
                cell = ws.cell(row=row_num, column=col_idx)
                if cell.value:
                    header_val = str(cell.value).strip().lower()
                    if "from" in header_val and from_col is None:
                        from_col = col_idx
                        header_row = row_num
                    elif "to" in header_val and to_col is None and col_idx != from_col:
                        if "tb" not in header_val and "no" not in header_val:
                            to_col = col_idx
                            if header_row is None:
                                header_row = row_num
                    elif ("final" in header_val and ("revis" in header_val or "revisi" in header_val)) and final_revision_col is None:
                        final_revision_col = col_idx
                        if header_row is None:
                            header_row = row_num
    
    if not (from_col and to_col and final_revision_col):
        if debug:
            print(f"  [DC Lookup] Columns not found. From={from_col}, To={to_col}, Final={final_revision_col}, HeaderRow={header_row}", file=sys.stderr)
            if header_row:
                print(f"  [DC Lookup] Header row {header_row} values:", file=sys.stderr)
                for c in range(1, min(ws.max_column + 1, 10)):
                    val = ws.cell(row=header_row, column=c).value
                    if val:
                        print(f"    Col {c}: {val}", file=sys.stderr)
        return None
    
    if debug:
        print(f"  [DC Lookup] Found columns: From={from_col}, To={to_col}, Final={final_revision_col}, HeaderRow={header_row}", file=sys.stderr)
    
    # Normalize time strings for comparison
    from_time_norm = normalize_time_str(from_time_str)
    to_time_norm = normalize_time_str(to_time_str)
    
    if debug:
        print(f"  [DC Lookup] Looking for time range: {from_time_norm} - {to_time_norm}", file=sys.stderr)
    
    # Search for matching time range
    data_start = header_row + 1
    matches_checked = 0
    max_rows_to_check = min(ws.max_row + 1, data_start + 200)  # Limit search to first 200 rows
    
    if debug:
        print(f"  [DC Lookup] Searching rows {data_start} to {max_rows_to_check-1} for time range", file=sys.stderr)
    
    for row_num in range(data_start, max_rows_to_check):
        from_cell = ws.cell(row=row_num, column=from_col)
        to_cell = ws.cell(row=row_num, column=to_col)
        
        if from_cell.value is None or to_cell.value is None:
            continue
        
        # Format and normalize cell values for comparison
        from_val_raw = format_value(from_cell.value)
        to_val_raw = format_value(to_cell.value)
        from_val = normalize_time_str(from_val_raw)
        to_val = normalize_time_str(to_val_raw)
        
        matches_checked += 1
        
        if debug and matches_checked <= 5:
            print(f"  [DC Lookup] Row {row_num}: '{from_val_raw}' -> '{from_val}' / '{to_val_raw}' -> '{to_val}'", file=sys.stderr)
        
        # Match time range
        if from_val == from_time_norm and to_val == to_time_norm:
            dc_cell = ws.cell(row=row_num, column=final_revision_col)
            dc_value = dc_cell.value
            if debug:
                print(f"  [DC Lookup] ✓ MATCH! Row {row_num}: {from_val} - {to_val}, DC value: {dc_value}", file=sys.stderr)
            return dc_value
    
    if debug:
        print(f"  [DC Lookup] ✗ No match found. Checked {matches_checked} rows.", file=sys.stderr)
        print(f"  [DC Lookup] Looking for: '{from_time_norm}' - '{to_time_norm}'", file=sys.stderr)
        # Show a few sample rows for comparison
        print(f"  [DC Lookup] Sample rows checked:", file=sys.stderr)
        sample_count = 0
        for row_num in range(data_start, min(max_rows_to_check, data_start + 10)):
            from_cell = ws.cell(row=row_num, column=from_col)
            to_cell = ws.cell(row=row_num, column=to_col)
            if from_cell.value is not None and to_cell.value is not None:
                from_val = normalize_time_str(format_value(from_cell.value))
                to_val = normalize_time_str(format_value(to_cell.value))
                print(f"    Row {row_num}: {from_val} - {to_val}", file=sys.stderr)
                sample_count += 1
                if sample_count >= 5:
                    break
    
    return None


def find_column_by_name(ws, column_name, max_header_rows=10):
    """
    Search for a column in the sheet by header name (case-insensitive, partial match).
    Scans first max_header_rows for header row.
    Returns (1-based column index, header_row) or (None, None).
    """
    target = column_name.strip().lower()
    if not target:
        return None, None
    
    for row_num in range(1, min(max_header_rows + 1, ws.max_row + 1)):
        for col_idx, cell in enumerate(ws[row_num], start=1):
            if cell.value is None:
                continue
            val = str(cell.value).strip().lower()
            # Exact match first
            if val == target:
                return col_idx, row_num
            # Partial match (target contained in header or header contained in target)
            if target in val or val in target:
                return col_idx, row_num
    
    return None, None


def find_matching_rows(ws, station_col_idx, station_name, header_row):
    """
    Find all rows where the station column matches the given station name.
    Returns list of (row_num, row_data) tuples.
    """
    target = station_name.strip()
    matches = []
    data_start = (header_row or 1) + 1
    
    for row_num in range(data_start, ws.max_row + 1):
        cell = ws.cell(row=row_num, column=station_col_idx)
        if cell.value is None:
            continue
        
        cell_val = str(cell.value).strip()
        # Exact match
        if cell_val == target:
            row_data = [ws.cell(row=row_num, column=c).value 
                       for c in range(1, ws.max_column + 1)]
            matches.append((row_num, row_data))
        # Case-insensitive match
        elif cell_val.upper() == target.upper():
            row_data = [ws.cell(row=row_num, column=c).value 
                       for c in range(1, ws.max_column + 1)]
            matches.append((row_num, row_data))
        # Partial match (station name contained in cell value)
        elif target.upper() in cell_val.upper():
            row_data = [ws.cell(row=row_num, column=c).value 
                       for c in range(1, ws.max_column + 1)]
            matches.append((row_num, row_data))
    
    return matches


def main():
    parser = argparse.ArgumentParser(
        description="Find rows in XLSX file where 'Name of the station' column matches given station name."
    )
    parser.add_argument(
        "xlsx_path",
        type=Path,
        help="Path to the XLSX file",
    )
    parser.add_argument(
        "station_name",
        help="Station name to search for (e.g., HINDUJA)",
    )
    parser.add_argument(
        "--sheet",
        help="Sheet name to search in (default: first/active sheet)",
        default=None,
    )
    parser.add_argument(
        "--column",
        help="Column name to search (default: searches for 'Name of the station')",
        default="Name of the station",
    )
    parser.add_argument(
        "--header-rows",
        type=int,
        default=10,
        help="Max header rows to scan for column name (default: 10)",
    )
    parser.add_argument(
        "--data-only",
        action="store_true",
        help="Read with data_only=True (evaluated values, not formulas)",
    )
    parser.add_argument(
        "--max-columns",
        type=int,
        default=20,
        help="Max columns to display per row (default: 20)",
    )
    parser.add_argument(
        "--dc-file",
        type=Path,
        help="Path to DC Excel file with date-named sheets (e.g., '01.01.2026')",
        default=None,
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Enable verbose debug output for DC lookup",
    )
    
    args = parser.parse_args()
    
    xlsx_path = args.xlsx_path
    if not xlsx_path.is_file():
        print(f"Error: File not found: {xlsx_path}", file=sys.stderr)
        sys.exit(1)
    
    # Load DC workbook if provided
    dc_wb = None
    if args.dc_file:
        # Try to resolve the path - check multiple locations
        dc_path = args.dc_file
        dc_path_resolved = None
        
        # List of paths to try
        paths_to_try = []
        
        if dc_path.is_absolute():
            # Absolute path - try as-is
            paths_to_try.append(dc_path)
        else:
            # Relative paths - try multiple locations
            # 1. Relative to input file directory
            paths_to_try.append(xlsx_path.parent / dc_path)
            # 2. Relative to current working directory
            paths_to_try.append(Path.cwd() / dc_path)
            # 3. Just the filename in input file directory
            if dc_path.name != str(dc_path):
                paths_to_try.append(xlsx_path.parent / dc_path.name)
            # 4. Just the filename in current directory
            if dc_path.name != str(dc_path):
                paths_to_try.append(Path.cwd() / dc_path.name)
        
        # Try each path
        for path_attempt in paths_to_try:
            if path_attempt.is_file():
                dc_path_resolved = path_attempt
                break
        
        if dc_path_resolved and dc_path_resolved.is_file():
            try:
                dc_wb = openpyxl.load_workbook(dc_path_resolved, read_only=True, data_only=True)
                print(f"DC file loaded: {dc_path_resolved}")
                print(f"  Available sheets ({len(dc_wb.sheetnames)} total): {', '.join(dc_wb.sheetnames[:10])}{'...' if len(dc_wb.sheetnames) > 10 else ''}")
                if len(dc_wb.sheetnames) > 10:
                    print(f"  ... and {len(dc_wb.sheetnames) - 10} more sheets")
            except Exception as e:
                print(f"Error loading DC file '{dc_path_resolved}': {e}", file=sys.stderr)
                print("Continuing without DC values...", file=sys.stderr)
        else:
            print(f"Warning: DC file not found: {args.dc_file}", file=sys.stderr)
            print(f"  Searched in:", file=sys.stderr)
            for path_attempt in paths_to_try:
                exists = "✓" if path_attempt.exists() else "✗"
                print(f"    {exists} {path_attempt}", file=sys.stderr)
            
            # Try to find similar files
            print(f"\n  Looking for similar files...", file=sys.stderr)
            search_dirs = [xlsx_path.parent, Path.cwd()]
            dc_filename_lower = str(args.dc_file).lower()
            found_similar = []
            
            for search_dir in search_dirs:
                if search_dir.exists() and search_dir.is_dir():
                    try:
                        for file in search_dir.glob("*.xlsx"):
                            if dc_filename_lower in file.name.lower() or file.name.lower() in dc_filename_lower:
                                found_similar.append(file)
                    except:
                        pass
            
            if found_similar:
                print(f"  Found similar files:", file=sys.stderr)
                for f in found_similar[:5]:
                    print(f"    - {f}", file=sys.stderr)
                if len(found_similar) > 5:
                    print(f"    ... and {len(found_similar) - 5} more", file=sys.stderr)
            
            print("Continuing without DC values...", file=sys.stderr)
    
    # Load workbook
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=args.data_only)
    
    # Select sheet
    if args.sheet:
        # Search for sheet by name
        sheet_found = None
        target = args.sheet.strip().lower()
        for name in wb.sheetnames:
            if name.strip().lower() == target or target in name.strip().lower():
                sheet_found = name
                break
        if sheet_found is None:
            print(f"Error: No sheet matching '{args.sheet}' in {xlsx_path.name}", file=sys.stderr)
            print(f"Available sheets: {', '.join(wb.sheetnames)}", file=sys.stderr)
            wb.close()
            sys.exit(1)
        ws = wb[sheet_found]
    else:
        ws = wb.active
    
    print(f"File: {xlsx_path.name}")
    print(f"Sheet: {ws.title}")
    print(f"Searching for station: '{args.station_name}'")
    print(f"Column to search: '{args.column}'")
    print()
    
    # Find column
    col_idx, header_row = find_column_by_name(ws, args.column, max_header_rows=args.header_rows)
    if col_idx is None:
        print(f"Error: No column matching '{args.column}' found in sheet '{ws.title}'", file=sys.stderr)
        print(f"Searched first {args.header_rows} rows.", file=sys.stderr)
        print(f"\nFirst row sample:", file=sys.stderr)
        first_row = [str(cell.value)[:30] if cell.value else "" for cell in ws[1][:15]]
        print(f"  {first_row}", file=sys.stderr)
        wb.close()
        sys.exit(1)
    
    col_letter = openpyxl.utils.get_column_letter(col_idx)
    print(f"Found column '{args.column}' at column {col_letter} (index {col_idx}), header row {header_row}")
    print()
    
    # Find matching rows
    matches = find_matching_rows(ws, col_idx, args.station_name, header_row)
    
    # Try to find "From Time" and "To Time" columns for 15-minute extraction
    from_time_col = None
    to_time_col = None
    date_col = None
    
    for col_idx_header in range(1, ws.max_column + 1):
        header_cell = ws.cell(row=header_row, column=col_idx_header)
        if header_cell.value:
            header_val = str(header_cell.value).strip().lower()
            if "from" in header_val and "time" in header_val:
                from_time_col = col_idx_header
                print(f"Found 'From Time' column at column {openpyxl.utils.get_column_letter(col_idx_header)}")
            elif "to" in header_val and "time" in header_val:
                to_time_col = col_idx_header
                print(f"Found 'To Time' column at column {openpyxl.utils.get_column_letter(col_idx_header)}")
            elif "date" in header_val:
                date_col = col_idx_header
                print(f"Found 'Date' column at column {openpyxl.utils.get_column_letter(col_idx_header)}")
    
    if not date_col:
        print("Warning: Date column not found. DC lookup will not work.", file=sys.stderr)
    if not from_time_col or not to_time_col:
        print("Warning: From/To Time columns not found. Time slots will not be generated.", file=sys.stderr)
    
    if not matches:
        print(f"No rows found where '{args.column}' = '{args.station_name}'")
        wb.close()
        sys.exit(0)
    
    print(f"Found {len(matches)} matching row(s):")
    print("=" * 100)
    
    # Get header row for display
    headers = []
    for col_idx_header in range(1, min(ws.max_column + 1, args.max_columns + 1)):
        header_cell = ws.cell(row=header_row, column=col_idx_header)
        headers.append(str(header_cell.value)[:25] if header_cell.value else f"Col{col_idx_header}")
    
    # Print header
    print(" | ".join(f"{h:<25}" for h in headers))
    print("-" * 100)
    
    # Print matching rows (without row numbers) and extract 15-minute ranges
    prev_entry = None
    for idx, (row_num, row_data) in enumerate(matches, 1):
        row_display = [format_value(v)[:25] for v in row_data[:args.max_columns]]
        print(" | ".join(f"{val:<25}" for val in row_display))
        if len(row_data) > args.max_columns:
            print(f"... ({len(row_data)} columns total, showing first {args.max_columns})")
        
        # Extract and display 15-minute intervals if time columns exist
        if from_time_col and to_time_col and from_time_col <= len(row_data) and to_time_col <= len(row_data):
            from_time_val = row_data[from_time_col - 1] if from_time_col > 0 else None
            to_time_val = row_data[to_time_col - 1] if to_time_col > 0 else None
            date_val = row_data[date_col - 1] if date_col and date_col > 0 and date_col <= len(row_data) else None
            
            if from_time_val is not None and to_time_val is not None:
                slots = slots_15min(from_time_val, to_time_val)
                if slots:
                    date_str = format_value(date_val) if date_val else ""
                    fr_str = format_value(from_time_val)
                    to_str = format_value(to_time_val)
                    
                    print(f"\n  Row {idx} - 15-minute intervals ({len(slots)} slots):")
                    if date_str:
                        print(f"    Date: {date_str}")
                    print(f"    From: {fr_str} | To: {to_str}")
                    
                    # Calculate gap from previous entry
                    if prev_entry:
                        prev_date_str, prev_to_str, prev_to_min = prev_entry
                        curr_from_min = parse_time_str(fr_str)
                        
                        if curr_from_min is not None and prev_to_min is not None:
                            if date_str == prev_date_str:
                                gap_minutes = curr_from_min - prev_to_min
                                if gap_minutes < 0:
                                    gap_minutes += 24 * 60
                            else:
                                if prev_to_min == 0 and curr_from_min == 0:
                                    gap_minutes = 0
                                else:
                                    gap_to_midnight = (24 * 60) - prev_to_min if prev_to_min > 0 else 0
                                    gap_from_midnight = curr_from_min
                                    gap_minutes = gap_to_midnight + gap_from_midnight
                            
                            gap_slots = gap_minutes // 15
                            if gap_slots > 0:
                                print(f"    [Gap: {gap_slots} x 15-min slots ({gap_minutes} minutes) from {prev_to_str} to {fr_str}]")
                            elif gap_slots == 0:
                                print(f"    [No gap: continuous from {prev_to_str} to {fr_str}]")
                    
                    # Show first few and last few slots if many
                    if len(slots) <= 10:
                        for slot_from, slot_to in slots:
                            print(f"      {slot_from} - {slot_to}")
                    else:
                        for slot_from, slot_to in slots[:3]:
                            print(f"      {slot_from} - {slot_to}")
                        print(f"      ... ({len(slots) - 6} more slots) ...")
                        for slot_from, slot_to in slots[-3:]:
                            print(f"      {slot_from} - {slot_to}")
                    
                    # Store for next iteration
                    to_min = parse_time_str(to_str)
                    if to_min is not None:
                        prev_entry = (date_str, to_str, to_min)
                    print()
    
    # Create output Excel file
    output_wb = openpyxl.Workbook()
    output_sheet = output_wb.active
    output_sheet.title = "Time Intervals"
    
    # Define styles
    header_font = Font(bold=True, size=11)
    center_align = Alignment(horizontal='center', vertical='center')
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Headers
    output_sheet['A1'] = 'Date'
    output_sheet['B1'] = 'From'
    output_sheet['C1'] = 'To'
    output_sheet['D1'] = 'DC (MW)'
    
    # Apply header styles
    for col in ['A1', 'B1', 'C1', 'D1']:
        cell = output_sheet[col]
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
    
    # Populate data rows
    row_idx = 2
    dc_found_count = 0
    dc_not_found_count = 0
    
    for idx, (row_num, row_data) in enumerate(matches, 1):
        if from_time_col and to_time_col and from_time_col <= len(row_data) and to_time_col <= len(row_data):
            from_time_val = row_data[from_time_col - 1] if from_time_col > 0 else None
            to_time_val = row_data[to_time_col - 1] if to_time_col > 0 else None
            date_val = row_data[date_col - 1] if date_col and date_col > 0 and date_col <= len(row_data) else None
            
            if from_time_val is not None and to_time_val is not None:
                slots = slots_15min(from_time_val, to_time_val)
                if slots:
                    date_str = format_value(date_val) if date_val else ""
                    
                    for slot_idx, (slot_from, slot_to) in enumerate(slots):
                        # Write date in first slot of each time range group
                        if slot_idx == 0 and date_str:
                            output_sheet.cell(row=row_idx, column=1).value = date_str
                        else:
                            output_sheet.cell(row=row_idx, column=1).value = ""  # Empty for subsequent slots in same range
                        
                        output_sheet.cell(row=row_idx, column=2).value = slot_from
                        output_sheet.cell(row=row_idx, column=3).value = slot_to
                        
                        # Lookup DC value if DC file is provided
                        dc_value = None
                        if dc_wb:
                            if date_str:
                                sheet_name = convert_date_to_sheet_format(date_str)
                                if sheet_name:
                                    # Enable debug for first few lookups and first slot of each unique date
                                    # Or if verbose flag is set
                                    debug_lookup = args.verbose or (row_idx <= 15) or (slot_idx == 0 and dc_not_found_count < 10)
                                    if debug_lookup:
                                        print(f"\n  [Row {row_idx}] Looking up DC for: Date={date_str} -> Sheet='{sheet_name}', Time={slot_from}-{slot_to}", file=sys.stderr)
                                    dc_value = find_dc_value(dc_wb, sheet_name, slot_from, slot_to, debug=debug_lookup)
                                    if dc_value is not None:
                                        dc_found_count += 1
                                        if debug_lookup:
                                            print(f"  [Row {row_idx}] ✓ DC value found: {dc_value}", file=sys.stderr)
                                    else:
                                        dc_not_found_count += 1
                                        if debug_lookup:
                                            print(f"  [Row {row_idx}] ✗ No DC value found", file=sys.stderr)
                                else:
                                    dc_not_found_count += 1
                                    if row_idx <= 10:
                                        print(f"  [Row {row_idx}] ✗ Could not convert date '{date_str}' to sheet format", file=sys.stderr)
                            else:
                                # No date available for this slot - this shouldn't happen for first slot
                                if slot_idx == 0:
                                    print(f"  [Row {row_idx}] ⚠ Warning: No date available for DC lookup (first slot)", file=sys.stderr)
                        else:
                            # DC file not provided or not loaded
                            if row_idx == 2:
                                print(f"  [Row {row_idx}] ⚠ Warning: DC workbook not available for lookup", file=sys.stderr)
                        
                        output_sheet.cell(row=row_idx, column=4).value = dc_value if dc_value is not None else ""
                        
                        # Apply borders
                        for col in range(1, 5):
                            output_sheet.cell(row=row_idx, column=col).border = thin_border
                        
                        row_idx += 1
    
    # Adjust column widths
    output_sheet.column_dimensions['A'].width = 15  # Date
    output_sheet.column_dimensions['B'].width = 10  # From
    output_sheet.column_dimensions['C'].width = 10  # To
    output_sheet.column_dimensions['D'].width = 12  # DC (MW)
    
    # Generate output filename with station name and timestamp (human-readable format with AM/PM)
    # Best practice: Use dashes for all separators (safe on all OS, readable)
    now = datetime.now()
    date_part = now.strftime("%d-%b-%Y")
    # Format time as "2-30-25-PM" (dashes instead of colons, no spaces, no leading zero on hour)
    hour = now.hour % 12
    if hour == 0:
        hour = 12
    time_part = f"{hour}-{now.minute:02d}-{now.second:02d}-{now.strftime('%p')}"
    timestamp = f"{date_part}_{time_part}"
    station_safe = args.station_name.replace(" ", "_").replace("/", "_")
    output_filename = f"{station_safe}_{timestamp}.xlsx"
    
    # Create output directory if it doesn't exist
    output_dir = xlsx_path.parent / "output"
    output_dir.mkdir(exist_ok=True)
    
    output_path = output_dir / output_filename
    
    output_wb.save(output_path)
    print(f"\nOutput file created: {output_path}")
    
    if dc_wb:
        print(f"\nDC Lookup Summary:")
        print(f"  DC values found: {dc_found_count}")
        print(f"  DC values not found: {dc_not_found_count}")
        if dc_not_found_count > 0 and dc_found_count == 0:
            print(f"\n  Warning: No DC values were found. Please check:")
            print(f"    - Date format conversion (instructions date -> DC sheet name)")
            print(f"    - Sheet names in DC file match date format")
            print(f"    - Time format matches between files")
            print(f"    - Column headers in DC file ('From', 'To', 'Final Revison')")
    
    wb.close()
    if dc_wb:
        dc_wb.close()


if __name__ == "__main__":
    main()
