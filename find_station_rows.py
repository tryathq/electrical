#!/usr/bin/env python3
"""
Find rows in XLSX file where "Name of the station" column matches given station name.

Usage:
  python find_station_rows.py <xlsx_path> <station_name> [options]

Example:
  python find_station_rows.py "input/Back_Down_Instructions.xlsx" HINDUJA
  python find_station_rows.py "input/jan 2026.xlsx" HINDUJA --sheet HNPCL
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
    
    args = parser.parse_args()
    
    xlsx_path = args.xlsx_path
    if not xlsx_path.is_file():
        print(f"Error: File not found: {xlsx_path}", file=sys.stderr)
        sys.exit(1)
    
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
            elif "to" in header_val and "time" in header_val:
                to_time_col = col_idx_header
            elif "date" in header_val:
                date_col = col_idx_header
    
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
    
    # Apply header styles
    for col in ['A1', 'B1', 'C1']:
        cell = output_sheet[col]
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
    
    # Populate data rows
    row_idx = 2
    prev_date = None
    
    for idx, (row_num, row_data) in enumerate(matches, 1):
        if from_time_col and to_time_col and from_time_col <= len(row_data) and to_time_col <= len(row_data):
            from_time_val = row_data[from_time_col - 1] if from_time_col > 0 else None
            to_time_val = row_data[to_time_col - 1] if to_time_col > 0 else None
            date_val = row_data[date_col - 1] if date_col and date_col > 0 and date_col <= len(row_data) else None
            
            if from_time_val is not None and to_time_val is not None:
                slots = slots_15min(from_time_val, to_time_val)
                if slots:
                    date_str = format_value(date_val) if date_val else ""
                    
                    for slot_from, slot_to in slots:
                        # Only write date in first row of each date group
                        if date_str and date_str != prev_date:
                            output_sheet.cell(row=row_idx, column=1).value = date_str
                            prev_date = date_str
                        else:
                            output_sheet.cell(row=row_idx, column=1).value = ""  # Empty for same date
                        
                        output_sheet.cell(row=row_idx, column=2).value = slot_from
                        output_sheet.cell(row=row_idx, column=3).value = slot_to
                        
                        # Apply borders
                        for col in range(1, 4):
                            output_sheet.cell(row=row_idx, column=col).border = thin_border
                        
                        row_idx += 1
    
    # Adjust column widths
    output_sheet.column_dimensions['A'].width = 15  # Date
    output_sheet.column_dimensions['B'].width = 10  # From
    output_sheet.column_dimensions['C'].width = 10  # To
    
    # Generate output filename with station name and timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    station_safe = args.station_name.replace(" ", "_").replace("/", "_")
    output_filename = f"{station_safe}_{timestamp}.xlsx"
    output_path = xlsx_path.parent / output_filename
    
    output_wb.save(output_path)
    print(f"\nOutput file created: {output_path}")
    
    wb.close()


if __name__ == "__main__":
    main()
