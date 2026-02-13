#!/usr/bin/env python3
"""
Streamlit Desktop App for Find Station Rows
Converts the command-line tool into a user-friendly GUI
"""

import streamlit as st
import sys
import re
from pathlib import Path
from datetime import datetime, date, time
import tempfile
import os
import pandas as pd

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    st.error("❌ Missing dependency: openpyxl. Please install with: pip install openpyxl")
    st.stop()

# Try to import streamlit-aggrid for advanced table features
try:
    from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode
    AGGrid_AVAILABLE = True
except ImportError:
    AGGrid_AVAILABLE = False

# Import functions from the original script
import sys
from pathlib import Path

# Add current directory to path to import find_station_rows module
sys.path.insert(0, str(Path(__file__).parent))

# Import the module (it will execute, but we'll use its functions)
try:
    import find_station_rows as fsr
    # Get all the functions we need
    format_value = fsr.format_value
    time_to_minutes = fsr.time_to_minutes
    floor_to_15 = fsr.floor_to_15
    minutes_to_time_str = fsr.minutes_to_time_str
    slots_15min = fsr.slots_15min
    parse_time_str = fsr.parse_time_str
    convert_date_to_sheet_format = fsr.convert_date_to_sheet_format
    normalize_time_str = fsr.normalize_time_str
    convert_date_for_bd_filename = fsr.convert_date_for_bd_filename
    find_bd_file = fsr.find_bd_file
    SCADALookupCache = fsr.SCADALookupCache
    find_scada_value = fsr.find_scada_value
    find_dc_value = fsr.find_dc_value
    find_column_by_name = fsr.find_column_by_name
    find_matching_rows = fsr.find_matching_rows
except ImportError as e:
    st.error(f"Failed to import find_station_rows module: {e}")
    st.stop()

# Page config
st.set_page_config(
    page_title="Report",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Title will be updated after processing with date range
if 'report_title' not in st.session_state:
    st.session_state.report_title = "⚡ REPORT"
    st.session_state.report_subtitle = "Generate electrical station data reports with time intervals"

# Sidebar for inputs
with st.sidebar:
    st.header("📋 Input Files")
    
    # Instructions file upload
    instructions_file = st.file_uploader(
        "Instructions Excel File",
        type=['xlsx', 'xls'],
        help="Upload the instructions XLSX file",
        key="instructions_file_upload"
    )
    
    # Sheet name (optional) - removed from UI, defaults to active sheet
    sheet_name = ""
    
    # Column name (read-only)
    column_name = st.text_input(
        "Column Name",
        value="Name of the station",
        help="Column header to search for station name",
        disabled=True
    )
    
    # Extract unique station names from file
    station_names = []
    station_name = None
    
    if instructions_file is not None:
        # Use session state to cache station names per file
        file_key = f"{instructions_file.name}_{sheet_name}_{column_name}"
        
        if 'station_names_cache' not in st.session_state:
            st.session_state.station_names_cache = {}
        
        # Always extract dates for title, even if station names are cached
        date_cache_key = f"{instructions_file.name}_{sheet_name}_dates"
        if 'date_range_cache' not in st.session_state:
            st.session_state.date_range_cache = {}
        
        if date_cache_key not in st.session_state.date_range_cache or file_key not in st.session_state.station_names_cache:
            with st.spinner("Extracting station names and dates from file..."):
                tmp_path = None
                try:
                    # Reset file pointer
                    instructions_file.seek(0)
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                        tmp_file.write(instructions_file.getbuffer())
                        tmp_path = tmp_file.name
                    
                    wb_temp = openpyxl.load_workbook(tmp_path, read_only=True, data_only=True)
                    
                    # Select sheet
                    if sheet_name:
                        sheet_found = None
                        target = sheet_name.strip().lower()
                        for name in wb_temp.sheetnames:
                            if name.strip().lower() == target or target in name.strip().lower():
                                sheet_found = name
                                break
                        if sheet_found:
                            ws_temp = wb_temp[sheet_found]
                        else:
                            ws_temp = wb_temp.active
                    else:
                        ws_temp = wb_temp.active
                    
                    # Find column
                    col_idx_temp, header_row_temp = find_column_by_name(ws_temp, column_name, max_header_rows=10)
                    
                    # Find date column
                    date_col_temp = None
                    for col_idx_header in range(1, min(ws_temp.max_column + 1, 50)):
                        header_cell = ws_temp.cell(row=header_row_temp, column=col_idx_header)
                        if header_cell.value:
                            header_val = str(header_cell.value).strip().lower()
                            if "date" in header_val:
                                date_col_temp = col_idx_header
                                break
                    
                    if col_idx_temp:
                        # Extract unique station names
                        unique_stations = set()
                        data_start = (header_row_temp or 1) + 1
                        max_rows_to_check = min(ws_temp.max_row + 1, data_start + 10000)  # Limit to 10k rows
                        
                        # Extract dates for title
                        dates_found = []
                        for row_num in range(data_start, max_rows_to_check):
                            cell = ws_temp.cell(row=row_num, column=col_idx_temp)
                            if cell.value:
                                station_val = str(cell.value).strip()
                                if station_val:
                                    unique_stations.add(station_val)
                            
                            # Extract date if date column found
                            if date_col_temp:
                                date_cell = ws_temp.cell(row=row_num, column=date_col_temp)
                                if date_cell.value:
                                    date_val = format_value(date_cell.value)
                                    if date_val:
                                        dates_found.append(date_val)
                        
                        station_names = sorted(list(unique_stations))
                        st.session_state.station_names_cache[file_key] = station_names
                        
                        # Extract and update date range for title
                        if dates_found:
                            parsed_dates = []
                            for d in dates_found:
                                try:
                                    for fmt in ["%d-%b-%Y", "%d-%b-%y", "%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d"]:
                                        try:
                                            parsed_dates.append((datetime.strptime(d, fmt), d))
                                            break
                                        except ValueError:
                                            continue
                                except:
                                    pass
                            
                            if parsed_dates:
                                parsed_dates.sort(key=lambda x: x[0])
                                report_from_date = parsed_dates[0][1]
                                report_to_date = parsed_dates[-1][1]
                                if report_from_date == report_to_date:
                                    title_str = f"⚡ REPORT FROM {report_from_date}"
                                else:
                                    title_str = f"⚡ REPORT FROM {report_from_date} TO {report_to_date}"
                                st.session_state.report_title = title_str
                                st.session_state.date_range_cache[date_cache_key] = title_str
                            else:
                                dates_sorted = sorted(set(dates_found))
                                if len(dates_sorted) == 1:
                                    title_str = f"⚡ REPORT FROM {dates_sorted[0]}"
                                elif len(dates_sorted) > 1:
                                    title_str = f"⚡ REPORT FROM {dates_sorted[0]} TO {dates_sorted[-1]}"
                                else:
                                    title_str = "⚡ REPORT"
                                st.session_state.report_title = title_str
                                st.session_state.date_range_cache[date_cache_key] = title_str
                        else:
                            st.session_state.report_title = "⚡ REPORT"
                            st.session_state.date_range_cache[date_cache_key] = "⚡ REPORT"
                    else:
                        st.session_state.station_names_cache[file_key] = []
                        st.session_state.date_range_cache[date_cache_key] = "⚡ REPORT"
                    
                    wb_temp.close()
                except Exception as e:
                    st.warning(f"Could not extract station names: {e}")
                    st.session_state.station_names_cache[file_key] = []
                    st.session_state.date_range_cache[date_cache_key] = "⚡ REPORT"
                finally:
                    if tmp_path and os.path.exists(tmp_path):
                        try:
                            os.unlink(tmp_path)
                        except:
                            pass
        else:
            station_names = st.session_state.station_names_cache[file_key]
            # Restore title from cache
            if date_cache_key in st.session_state.date_range_cache:
                st.session_state.report_title = st.session_state.date_range_cache[date_cache_key]
        
        # Show dropdown (always selectbox, never editable text input)
        if station_names:
            station_name = st.selectbox(
                "Station Name",
                options=station_names,
                help=f"Select station name from the dropdown ({len(station_names)} stations found in file)",
                key="station_selectbox"
            )
            st.caption(f"✓ Found {len(station_names)} unique station(s)")
        else:
            # Show empty selectbox (not editable) if no stations found
            station_name = st.selectbox(
                "Station Name",
                options=[],
                help="No stations found in file. Please check the file and column name.",
                disabled=True,
                key="station_selectbox_empty"
            )
            if column_name:
                st.caption("⚠️ No stations found. Check if column name matches the file.")
    else:
        # No file uploaded yet, show text input
        station_name = st.text_input(
            "Station Name",
            value="",
            help="Upload instructions file to see dropdown, or enter station name manually"
        )
    
    st.divider()
    st.header("⚙️ Options")
    
    # DC file upload (mandatory)
    dc_file = st.file_uploader(
        "DC File",
        type=['xlsx', 'xls'],
        help="DC Excel file with date-named sheets (required)"
    )
    
    # BD folder (mandatory)
    bd_folder_path = st.text_input(
        "BD Folder Path",
        value="",
        help="Path to folder containing SCADA BD files (required)"
    )
    
    # BD sheet name (mandatory) - extract from BD file
    bd_sheet = ""
    bd_sheet_options = []
    
    if bd_folder_path and bd_folder_path.strip():
        # Try to find a BD file and extract sheet names
        bd_folder = Path(bd_folder_path.strip())
        if bd_folder.exists() and bd_folder.is_dir():
            # Find first Excel file in BD folder
            bd_files = list(bd_folder.glob("*.xlsx")) + list(bd_folder.glob("*.xls"))
            if bd_files:
                bd_file_path = bd_files[0]
                file_key_sheets = f"{bd_file_path.name}_sheets"
                
                if 'bd_sheets_cache' not in st.session_state:
                    st.session_state.bd_sheets_cache = {}
                
                if file_key_sheets not in st.session_state.bd_sheets_cache:
                    try:
                        with st.spinner("Extracting sheet names from BD file..."):
                            wb_bd = openpyxl.load_workbook(bd_file_path, read_only=True, data_only=True)
                            bd_sheet_options = wb_bd.sheetnames
                            st.session_state.bd_sheets_cache[file_key_sheets] = bd_sheet_options
                            wb_bd.close()
                    except Exception as e:
                        st.session_state.bd_sheets_cache[file_key_sheets] = []
                else:
                    bd_sheet_options = st.session_state.bd_sheets_cache[file_key_sheets]
    
    # Show BD sheet dropdown if options available
    if bd_sheet_options:
        bd_sheet = st.selectbox(
            "BD Sheet Name",
            options=bd_sheet_options,
            help="Select sheet name from BD file (extracted from BD folder)",
            key="bd_sheet_selectbox"
        )
        st.caption(f"✓ Found {len(bd_sheet_options)} sheet(s) in BD file")
    else:
        bd_sheet = st.text_input(
            "BD Sheet Name",
            value="",
            help="Sheet name in BD files (e.g., 'DATA-CMD') (required)",
            key="bd_sheet_text"
        )
        if bd_folder_path:
            st.caption("⚠️ Could not extract sheets. Check BD folder path.")
    
    # SCADA column (mandatory) - extract from BD file
    scada_column = None
    scada_column_options = []
    
    if bd_folder_path and bd_folder_path.strip() and bd_sheet and str(bd_sheet).strip():
        # Try to find a BD file and extract column names
        bd_folder = Path(bd_folder_path.strip())
        if bd_folder.exists() and bd_folder.is_dir():
            # Find first Excel file in BD folder
            bd_files = list(bd_folder.glob("*.xlsx")) + list(bd_folder.glob("*.xls"))
            if bd_files:
                bd_file_path = bd_files[0]
                file_key_cols = f"{bd_file_path.name}_{bd_sheet.strip()}_columns"
                
                if 'bd_columns_cache' not in st.session_state:
                    st.session_state.bd_columns_cache = {}
                
                if file_key_cols not in st.session_state.bd_columns_cache:
                    try:
                        with st.spinner("Extracting column names from BD file..."):
                            wb_bd = openpyxl.load_workbook(bd_file_path, read_only=True, data_only=True)
                            
                            # Find the specified sheet
                            sheet_found = None
                            target_sheet = bd_sheet.strip().lower()
                            for name in wb_bd.sheetnames:
                                if name.strip().lower() == target_sheet or target_sheet in name.strip().lower():
                                    sheet_found = name
                                    break
                            
                            if sheet_found:
                                ws_bd = wb_bd[sheet_found]
                                # Extract column names only from header row (typically row 1)
                                column_names = []
                                # Most Excel files have headers in row 1
                                header_row = 1
                                
                                # Extract all values from header row
                                for col_idx in range(1, min(ws_bd.max_column + 1, 200)):
                                    cell = ws_bd.cell(row=header_row, column=col_idx)
                                    if cell.value:
                                        col_name = str(cell.value).strip()
                                        if col_name:
                                            column_names.append(col_name)
                                
                                # If row 1 is empty or has very few values, try row 2
                                if len(column_names) < 2 and ws_bd.max_row >= 2:
                                    column_names = []
                                    header_row = 2
                                    for col_idx in range(1, min(ws_bd.max_column + 1, 200)):
                                        cell = ws_bd.cell(row=header_row, column=col_idx)
                                        if cell.value:
                                            col_name = str(cell.value).strip()
                                            if col_name:
                                                column_names.append(col_name)
                                
                                scada_column_options = sorted(list(set(column_names)))  # Remove duplicates, sort
                                st.session_state.bd_columns_cache[file_key_cols] = scada_column_options
                            else:
                                st.session_state.bd_columns_cache[file_key_cols] = []
                            
                            wb_bd.close()
                    except Exception as e:
                        st.session_state.bd_columns_cache[file_key_cols] = []
                else:
                    scada_column_options = st.session_state.bd_columns_cache[file_key_cols]
    
    # Show SCADA column dropdown if options available
    if scada_column_options:
        scada_column = st.selectbox(
            "SCADA Column Name",
            options=scada_column_options,
            help="Select column name from BD file (extracted from BD folder)",
            key="scada_column_selectbox"
        )
        st.caption(f"✓ Found {len(scada_column_options)} column(s) in BD file")
    else:
        scada_column = st.text_input(
            "SCADA Column Name",
            value="",
            help="Column header name in BD files (e.g., 'HNJA4_AG.STTN.X_BUS_GEN.MW') (required)",
            key="scada_column_text"
        )
        if bd_folder_path and bd_sheet:
            st.caption("⚠️ Could not extract columns. Check BD folder path and sheet name.")
    
    # Advanced options
    with st.expander("🔧 Advanced Options"):
        header_rows = st.number_input(
            "Max Header Rows",
            min_value=1,
            max_value=50,
            value=10,
            help="Maximum rows to scan for column headers"
        )
        data_only = st.checkbox(
            "Data Only Mode",
            value=False,
            help="Read with data_only=True (evaluated values, not formulas)"
        )
        verbose = st.checkbox(
            "Verbose Output",
            value=False,
            help="Enable verbose debug output"
        )

# Display title after sidebar processing (so it can be updated by file upload)
title_to_show = st.session_state.get('report_title', "⚡ REPORT")
st.title(title_to_show)
st.markdown(st.session_state.get('report_subtitle', "Generate electrical station data reports with time intervals"))

# Main content area
if instructions_file is None:
    st.info("👈 Please upload an Instructions Excel file in the sidebar to get started.")
    st.stop()

if not station_name or station_name.strip() == "":
    if 'station_names_cache' in st.session_state and len(st.session_state.station_names_cache) > 0:
        st.warning("⚠️ Please select a Station Name from the dropdown")
    else:
        st.warning("⚠️ Please enter or select a Station Name")
    st.stop()

# Check if DC file is provided (mandatory)
if dc_file is None:
    st.error("❌ DC File is required. Please upload a DC Excel file.")
    st.stop()

# Check if BD folder path is provided (mandatory)
if not bd_folder_path or bd_folder_path.strip() == "":
    st.error("❌ BD Folder Path is required. Please enter the path to the BD folder.")
    st.stop()

# Check if SCADA column is provided (mandatory)
if not scada_column or scada_column.strip() == "":
    st.error("❌ SCADA Column Name is required. Please enter the SCADA column name.")
    st.stop()

# Check if BD sheet name is provided (mandatory)
if not bd_sheet or bd_sheet.strip() == "":
    st.error("❌ BD Sheet Name is required. Please enter the BD sheet name.")
    st.stop()

# Generate button
if st.button("🚀 Generate", type="primary", use_container_width=True):
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # Save uploaded files to temp directory
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            
            # Save instructions file
            instructions_path = temp_path / instructions_file.name
            with open(instructions_path, "wb") as f:
                f.write(instructions_file.getbuffer())
            
            # Save DC file if provided
            dc_path = None
            if dc_file:
                dc_path = temp_path / dc_file.name
                with open(dc_path, "wb") as f:
                    f.write(dc_file.getbuffer())
            
            # Handle BD folder
            bd_folder = None
            if bd_folder_path:
                bd_folder = Path(bd_folder_path)
                if not bd_folder.exists():
                    # Try relative to instructions file location
                    if instructions_file.name:
                        # Try common locations
                        possible_paths = [
                            Path(bd_folder_path),
                            Path("data") / "BD",
                            Path("data") / bd_folder_path,
                        ]
                        for pp in possible_paths:
                            if pp.exists() and pp.is_dir():
                                bd_folder = pp
                                break
                if bd_folder and bd_folder.exists() and bd_folder.is_dir():
                    pass
                else:
                    bd_folder = None
                    if scada_column:
                        st.warning(f"⚠️ BD folder not found: {bd_folder_path}")
            
            progress_bar.progress(10)
            status_text.text("Loading instructions file...")
            
            # Load workbook
            wb = openpyxl.load_workbook(instructions_path, read_only=True, data_only=data_only)
            
            # Select sheet
            if sheet_name:
                sheet_found = None
                target = sheet_name.strip().lower()
                for name in wb.sheetnames:
                    if name.strip().lower() == target or target in name.strip().lower():
                        sheet_found = name
                        break
                if sheet_found is None:
                    st.error(f"❌ Sheet '{sheet_name}' not found in {instructions_file.name}")
                    st.write(f"Available sheets: {', '.join(wb.sheetnames)}")
                    st.stop()
                ws = wb[sheet_found]
            else:
                ws = wb.active
            
            progress_bar.progress(20)
            status_text.text("Finding station column...")
            
            # Find column
            col_idx, header_row = find_column_by_name(ws, column_name, max_header_rows=header_rows)
            if col_idx is None:
                st.error(f"❌ Column '{column_name}' not found in sheet '{ws.title}'")
                st.write(f"Searched first {header_rows} rows.")
                st.stop()
            
            progress_bar.progress(30)
            status_text.text("Finding matching rows...")
            
            # Find matching rows
            matches = find_matching_rows(ws, col_idx, station_name, header_row)
            
            if not matches:
                st.warning(f"⚠️ No rows found where '{column_name}' = '{station_name}'")
                wb.close()
                st.stop()
            
            st.success(f"✅ Found {len(matches)} matching row(s)")
            
            # Find time columns and date column
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
            
            progress_bar.progress(40)
            status_text.text("Creating output file...")
            
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
            
            pad = 0
            header_row_out, start_col, start_data_row = 1 + pad, 1 + pad, 2 + pad
            
            # Headers
            output_sheet.cell(row=header_row_out, column=start_col).value = 'Date'
            output_sheet.cell(row=header_row_out, column=start_col + 1).value = 'From'
            output_sheet.cell(row=header_row_out, column=start_col + 2).value = 'To'
            output_sheet.cell(row=header_row_out, column=start_col + 3).value = 'DC (MW)'
            output_sheet.cell(row=header_row_out, column=start_col + 4).value = 'As per SLDC Scada in MW'
            output_sheet.cell(row=header_row_out, column=start_col + 5).value = 'Diff (MW)'
            
            # Apply header styles
            for c in range(6):
                cell = output_sheet.cell(row=header_row_out, column=start_col + c)
                cell.font = header_font
                cell.alignment = center_align
                cell.border = thin_border
            
            progress_bar.progress(50)
            status_text.text("Initializing caches...")
            
            # Initialize SCADA cache
            scada_cache = None
            if bd_folder and scada_column:
                scada_cache = SCADALookupCache(bd_folder, scada_column, bd_sheet if bd_sheet else None)
            
            # Load DC workbook
            dc_wb = None
            if dc_path and dc_path.exists():
                dc_wb = openpyxl.load_workbook(dc_path, read_only=True, data_only=True)
            
            # Populate data rows
            row_idx = start_data_row
            dc_found_count = 0
            dc_not_found_count = 0
            scada_found_count = 0
            scada_not_found_count = 0
            
            total_slots = 0
            processed_slots = 0
            current_date = None
            previous_date_with_data = None
            date_start_row = None
            
            # Count total slots
            for idx, (row_num, row_data) in enumerate(matches, 1):
                if from_time_col and to_time_col and from_time_col <= len(row_data) and to_time_col <= len(row_data):
                    from_time_val = row_data[from_time_col - 1] if from_time_col > 0 else None
                    to_time_val = row_data[to_time_col - 1] if to_time_col > 0 else None
                    if from_time_val is not None and to_time_val is not None:
                        slots = slots_15min(from_time_val, to_time_val)
                        total_slots += len(slots) if slots else 0
            
            progress_bar.progress(60)
            status_text.text(f"Processing {len(matches)} time range(s) with {total_slots} total slots...")
            
            # Process matches
            for idx, (row_num, row_data) in enumerate(matches, 1):
                if from_time_col and to_time_col and from_time_col <= len(row_data) and to_time_col <= len(row_data):
                    from_time_val = row_data[from_time_col - 1] if from_time_col > 0 else None
                    to_time_val = row_data[to_time_col - 1] if to_time_col > 0 else None
                    date_val = row_data[date_col - 1] if date_col and date_col > 0 and date_col <= len(row_data) else None
                    
                    if from_time_val is not None and to_time_val is not None:
                        slots = slots_15min(from_time_val, to_time_val)
                        if slots:
                            date_str = format_value(date_val) if date_val else ""
                            
                            # Merge previous date's cells if date changed
                            if date_str and date_str != previous_date_with_data and previous_date_with_data is not None:
                                if date_start_row is not None and row_idx > date_start_row:
                                    date_col_letter = get_column_letter(start_col)
                                    output_sheet.merge_cells(f"{date_col_letter}{date_start_row}:{date_col_letter}{row_idx - 1}")
                                row_idx += 1
                                date_start_row = None
                            
                            if date_str and date_str != current_date:
                                current_date = date_str
                                previous_date_with_data = date_str
                                date_start_row = row_idx
                                # Track date for report title
                            
                            for slot_idx, (slot_from, slot_to) in enumerate(slots):
                                if slot_idx == 0 and date_str and date_start_row == row_idx:
                                    date_cell = output_sheet.cell(row=row_idx, column=start_col)
                                    date_cell.value = date_str
                                    date_cell.alignment = Alignment(horizontal='center', vertical='center')
                                
                                output_sheet.cell(row=row_idx, column=start_col + 1).value = slot_from
                                output_sheet.cell(row=row_idx, column=start_col + 2).value = slot_to
                                
                                # DC lookup
                                dc_value = None
                                if dc_wb and date_str:
                                    sheet_name_dc = convert_date_to_sheet_format(date_str)
                                    if sheet_name_dc:
                                        dc_value = find_dc_value(dc_wb, sheet_name_dc, slot_from, slot_to, debug=verbose)
                                        if dc_value is not None:
                                            dc_found_count += 1
                                        else:
                                            dc_not_found_count += 1
                                
                                output_sheet.cell(row=row_idx, column=start_col + 3).value = dc_value if dc_value is not None else ""
                                
                                # SCADA lookup
                                scada_value = None
                                if scada_cache and date_str:
                                    scada_value = find_scada_value(scada_cache, date_str, slot_from, debug=verbose, show_progress=False)
                                    if scada_value is not None:
                                        scada_found_count += 1
                                    else:
                                        scada_not_found_count += 1
                                    processed_slots += 1
                                
                                output_sheet.cell(row=row_idx, column=start_col + 4).value = scada_value if scada_value is not None else ""
                                
                                # Calculate difference
                                diff_value = None
                                if dc_value is not None and scada_value is not None:
                                    try:
                                        dc_num = float(dc_value) if isinstance(dc_value, (int, float, str)) and str(dc_value).strip() else None
                                        scada_num = float(scada_value) if isinstance(scada_value, (int, float, str)) and str(scada_value).strip() else None
                                        if dc_num is not None and scada_num is not None:
                                            diff_value = dc_num - scada_num
                                    except (ValueError, TypeError):
                                        pass
                                
                                output_sheet.cell(row=row_idx, column=start_col + 5).value = diff_value if diff_value is not None else ""
                                
                                # Apply borders
                                for c in range(6):
                                    output_sheet.cell(row=row_idx, column=start_col + c).border = thin_border
                                
                                row_idx += 1
                                
                                # Update progress
                                if total_slots > 0:
                                    progress = 60 + int(30 * processed_slots / total_slots)
                                    progress_bar.progress(min(progress, 90))
            
            # Merge last date
            if date_start_row is not None and row_idx > date_start_row:
                date_col_letter = get_column_letter(start_col)
                output_sheet.merge_cells(f"{date_col_letter}{date_start_row}:{date_col_letter}{row_idx - 1}")
            
            last_row = row_idx - 1
            last_content_col = start_col + 5
            
            # Freeze header
            output_sheet.freeze_panes = output_sheet.cell(row=start_data_row, column=start_col).coordinate
            
            # Hide gridlines
            output_sheet.sheet_view.showGridLines = False
            
            # Adjust column widths
            for i, w in enumerate([15, 10, 10, 12, 25, 12]):
                output_sheet.column_dimensions[get_column_letter(start_col + i)].width = w
            
            # Print area
            output_sheet.print_area = f'A1:{get_column_letter(last_content_col)}{last_row}'
            
            progress_bar.progress(95)
            status_text.text("Saving output file...")
            
            # Save to temp file
            output_filename = f"{station_name.replace(' ', '_').replace('/', '_')}_{datetime.now().strftime('%d-%b-%Y_%H-%M-%S-%p')}.xlsx"
            output_path = temp_path / output_filename
            output_wb.save(output_path)
            
            progress_bar.progress(100)
            status_text.text("✅ Processing complete!")
            
            # Close workbooks
            wb.close()
            if dc_wb:
                dc_wb.close()
            if scada_cache:
                scada_cache.close_all()
            
            # Title already updated from instructions file date range above
            
            # Display summary
            st.success("✅ Output file generated successfully!")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Rows", len(matches))
            with col2:
                st.metric("Total Slots", total_slots)
            with col3:
                st.metric("Output Rows", last_row - 1)
            
            if dc_wb:
                st.info(f"DC Lookups: {dc_found_count} found, {dc_not_found_count} not found")
            if scada_cache:
                st.info(f"SCADA Lookups: {scada_found_count} found, {scada_not_found_count} not found")
            
            # Load output data for display
            try:
                df_output = pd.read_excel(output_path, engine='openpyxl')
                
                # Create dynamic table title based on station and date range
                title_parts = ["calculation sheet for BD and non compliance of", station_name]
                
                # Extract date range from report title or use current dates
                report_title = st.session_state.get('report_title', "⚡ REPORT")
                if "FROM" in report_title and "TO" in report_title:
                    # Extract date part (e.g., "01-Jan-2026 TO 31-Jan-2026")
                    date_part = report_title.split("FROM")[1].strip() if "FROM" in report_title else ""
                    # Extract month/year for shorter format (e.g., "Jan 26")
                    try:
                        if "TO" in date_part:
                            from_date_str = date_part.split("TO")[0].strip()
                            to_date_str = date_part.split("TO")[1].strip()
                            # Parse and format as "Jan 26" if same month/year
                            try:
                                from_dt = datetime.strptime(from_date_str, "%d-%b-%Y")
                                to_dt = datetime.strptime(to_date_str, "%d-%b-%Y")
                                if from_dt.year == to_dt.year and from_dt.month == to_dt.month:
                                    date_suffix = f"for {from_dt.strftime('%b %y')}"
                                else:
                                    date_suffix = f"for {from_dt.strftime('%b %y')} to {to_dt.strftime('%b %y')}"
                            except:
                                date_suffix = f"for {date_part}"
                        else:
                            try:
                                from_dt = datetime.strptime(date_part, "%d-%b-%Y")
                                date_suffix = f"for {from_dt.strftime('%b %y')}"
                            except:
                                date_suffix = f"for {date_part}"
                        title_parts.append(date_suffix)
                    except:
                        pass
                elif "FROM" in report_title:
                    date_part = report_title.split("FROM")[1].strip()
                    try:
                        from_dt = datetime.strptime(date_part, "%d-%b-%Y")
                        title_parts.append(f"for {from_dt.strftime('%b %y')}")
                    except:
                        title_parts.append(f"for {date_part}")
                
                table_title = " ".join(title_parts)
                
                st.divider()
                st.header(f"📊 {table_title}")
                
                # Search and filter controls
                col_search1, col_search2 = st.columns([2, 1])
                with col_search1:
                    search_term = st.text_input(
                        "🔍 Search",
                        value="",
                        placeholder="Search in all columns...",
                        help="Filter rows by searching across all columns"
                    )
                with col_search2:
                    rows_per_page = st.selectbox(
                        "Rows per page",
                        options=[10, 25, 50, 100, 500],
                        index=1,  # Default to 25
                        help="Number of rows to display per page"
                    )
                
                # Apply search filter
                if search_term:
                    mask = df_output.astype(str).apply(
                        lambda x: x.str.contains(search_term, case=False, na=False)
                    ).any(axis=1)
                    df_filtered = df_output[mask].copy()
                else:
                    df_filtered = df_output.copy()
                
                # Pagination
                total_rows = len(df_filtered)
                total_pages = (total_rows + rows_per_page - 1) // rows_per_page if total_rows > 0 else 1
                
                if total_pages > 1:
                    page_num = st.number_input(
                        f"Page (1-{total_pages})",
                        min_value=1,
                        max_value=total_pages,
                        value=1,
                        step=1
                    )
                    start_idx = (page_num - 1) * rows_per_page
                    end_idx = start_idx + rows_per_page
                    df_display = df_filtered.iloc[start_idx:end_idx].copy()
                    st.caption(f"Showing rows {start_idx + 1} to {min(end_idx, total_rows)} of {total_rows} (Page {page_num}/{total_pages})")
                else:
                    df_display = df_filtered.copy()
                    st.caption(f"Showing all {total_rows} rows")
                
                # Display table with sorting
                if AGGrid_AVAILABLE:
                    # Use AgGrid for advanced features
                    gb = GridOptionsBuilder.from_dataframe(df_display)
                    gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=rows_per_page)
                    gb.configure_side_bar()
                    gb.configure_default_column(
                        sortable=True,
                        filterable=True,
                        resizable=True,
                        editable=False
                    )
                    gb.configure_selection('single')
                    gridOptions = gb.build()
                    
                    AgGrid(
                        df_display,
                        gridOptions=gridOptions,
                        height=400,
                        width='100%',
                        theme='streamlit',
                        update_mode=GridUpdateMode.NO_UPDATE,
                        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
                        allow_unsafe_jscode=True
                    )
                else:
                    # Use standard Streamlit dataframe with sorting
                    st.dataframe(
                        df_display,
                        use_container_width=True,
                        height=400,
                        hide_index=True
                    )
                    st.caption("💡 Tip: Install streamlit-aggrid for advanced filtering and sorting: pip install streamlit-aggrid")
                
            except Exception as e:
                st.warning(f"Could not display preview: {e}")
            
            # Download button
            st.divider()
            with open(output_path, "rb") as f:
                st.download_button(
                    label="📥 Download Output File",
                    data=f.read(),
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
    
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        if verbose:
            import traceback
            st.code(traceback.format_exc())
