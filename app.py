#!/usr/bin/env python3
"""
Streamlit Desktop App for Find Station Rows
Converts the command-line tool into a user-friendly GUI
"""

import streamlit as st
import sys
import uuid
from pathlib import Path
from datetime import datetime
import tempfile
import os
import json
import shutil
import pandas as pd

# Slots per batch for incremental table updates
PROCESSING_BATCH_SIZE = 5

# Reports persistence: folder and index path (on disk — persists after app close)
REPORTS_DIR = Path(__file__).resolve().parent / "reports"
REPORTS_INDEX_FILE = REPORTS_DIR / "reports_index.json"


def _reports_ensure_dir():
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)


def _reports_load_index():
    """Load list of persisted reports from disk (newest first)."""
    if not REPORTS_INDEX_FILE.exists():
        return []
    try:
        with open(REPORTS_INDEX_FILE, "r", encoding="utf-8") as f:
            entries = json.load(f)
        return sorted(entries, key=lambda e: e.get("run_at", ""), reverse=True)
    except Exception:
        return []


def _reports_append_entry(entry):
    """Append one report entry to the index and flush to disk so it persists after app close."""
    _reports_ensure_dir()
    entries = _reports_load_index()
    entries.insert(0, entry)
    with open(REPORTS_INDEX_FILE, "w", encoding="utf-8") as f:
        json.dump(entries, f, indent=2, ensure_ascii=False)
        f.flush()
        os.fsync(f.fileno())


def _reports_save_file(src_path: Path, filename: str) -> Path:
    """Copy report file to reports dir on disk; persists after app close."""
    _reports_ensure_dir()
    dest = REPORTS_DIR / filename
    shutil.copy2(src_path, dest)
    return dest

# Table height: ~35px per row + 40px header, min 200, max 500 (fits most screens)
def _table_height(row_count, min_h=200, max_h=500, row_px=35, header_px=40):
    return min(max(min_h, header_px + row_count * row_px), max_h)

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

# Add current directory to path to import find_station_rows module
sys.path.insert(0, str(Path(__file__).parent))

# Import the module (it will execute, but we'll use its functions)
try:
    import find_station_rows as fsr
    # Get the functions we need from find_station_rows
    format_value = fsr.format_value
    slots_15min = fsr.slots_15min
    convert_date_to_sheet_format = fsr.convert_date_to_sheet_format
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

def _url_reports_list():
    """Set URL to reports list only (?view=reports)."""
    if getattr(st, "query_params", None) is not None and hasattr(st.query_params, "from_dict"):
        st.query_params.from_dict({"view": "reports"})


def _url_report_file(filename):
    """Set URL to single report (?view=report&file=...)."""
    if getattr(st, "query_params", None) is not None and hasattr(st.query_params, "from_dict"):
        st.query_params.from_dict({"view": "report", "file": filename})


def _url_main():
    """Set URL to main page (no view/file params)."""
    if getattr(st, "query_params", None) is not None:
        if hasattr(st.query_params, "clear"):
            st.query_params.clear()
        elif hasattr(st.query_params, "from_dict"):
            st.query_params.from_dict({})


# Sync Reports view from URL (skip if user just navigated away so we don't re-apply stale URL)
if not st.session_state.pop("_url_go_main", None) and getattr(st, "query_params", None):
    qp = st.query_params
    if qp.get("view") == "report" and qp.get("file"):
        # Direct link to a report: ?view=report&file=filename.xlsx
        report_file = qp.get("file")
        for entry in _reports_load_index():
            if entry.get("filename") == report_file:
                st.session_state["reports_view_filename"] = report_file
                st.session_state["reports_view_entry"] = entry
                st.session_state["reports_view_from_list"] = True
                st.session_state.pop("view_mode", None)
                break
    elif qp.get("view") == "reports" and not st.session_state.get("reports_view_filename"):
        st.session_state["view_mode"] = "reports"
# Normalize URL when on main page: clear any stale view/file
if getattr(st, "query_params", None) and not st.session_state.get("view_mode") and not st.session_state.get("reports_view_filename"):
    qp = st.query_params
    if qp.get("view") or qp.get("file"):
        _url_main()

# Sidebar: Menu at top (big square buttons); then Home (generate form) or Reports (list of reports)
with st.sidebar:
    st.markdown('<div data-app-menu-row style="display:none" aria-hidden="true"></div>', unsafe_allow_html=True)
    view_mode = st.session_state.get("view_mode", "")
    _sidebar_home = not view_mode and not st.session_state.get("reports_view_filename")
    _on_report = view_mode == "reports" or st.session_state.get("reports_view_filename") or st.session_state.get("reports_view_active")
    col_h, col_r = st.columns(2)
    with col_h:
        if st.button("🏠 Home", key="sidebar_home", type="primary" if _sidebar_home else "secondary", use_container_width=True):
            for key in ("view_mode", "reports_view_filename", "reports_view_entry", "reports_view_active", "reports_view_from_list"):
                st.session_state.pop(key, None)
            st.session_state["_url_go_main"] = True
            _url_main()
            st.rerun()
    with col_r:
        if st.button("📂 Reports", key="sidebar_reports", type="primary" if _on_report else "secondary", use_container_width=True):
            st.session_state["view_mode"] = "reports"
            _url_reports_list()
            st.rerun()
    st.divider()
    if _sidebar_home:
        st.caption("**Home** — generate report")
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
        
        # Defaults (advanced options removed for now)
        header_rows = 10
        data_only = False
        verbose = False
    else:
        # Reports: show list of saved reports in sidebar; selecting one shows it on the right
        instructions_file = None
        station_name = ""
        dc_file = None
        bd_folder_path = ""
        scada_column = None
        bd_sheet = ""
        header_rows = 10
        data_only = False
        verbose = False
        st.caption("**Reports** — select a report")
        reports_list_sidebar = _reports_load_index()
        if not reports_list_sidebar:
            st.info("No saved reports yet. Go to **Home** to generate one.")
        else:
            st.caption(f"{len(reports_list_sidebar)} saved report(s)")
            _selected_report = st.session_state.get("reports_view_filename") or st.session_state.get("reports_view_active")
            for i, entry in enumerate(reports_list_sidebar):
                fn = entry.get("filename", "")
                station = entry.get("station", "")
                date_from = entry.get("date_from", "")
                date_to = entry.get("date_to", "")
                date_range = f"{date_from} → {date_to}" if date_to else (date_from or "—")
                label = f"{station} — {date_range}"
                run_at = entry.get("run_at", "")
                try:
                    dt = datetime.fromisoformat(run_at.replace("Z", "+00:00"))
                    generated_str = dt.strftime("%d %b %Y, %I:%M %p")
                except Exception:
                    generated_str = run_at if run_at else "—"
                # Two lines in one button (timestamp may show centered)
                label_with_time = f"{label}\n{generated_str}"
                _is_selected = (fn == _selected_report)
                c1, c2 = st.columns([5, 1])
                with c1:
                    if st.button(label_with_time, key=f"sidebar_rep_{i}_{fn}", type="primary" if _is_selected else "secondary", use_container_width=True, help="Show this report on the right"):
                        st.session_state["reports_view_filename"] = fn
                        st.session_state["reports_view_entry"] = entry
                        st.session_state["reports_view_from_list"] = True
                        _url_report_file(fn)
                        st.rerun()
                with c2:
                    report_path = REPORTS_DIR / fn
                    if report_path.exists():
                        with open(report_path, "rb") as f:
                            st.download_button("📥", data=f.read(), file_name=fn, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"sidebar_dl_{i}_{fn}")

# Global CSS for sidebar menu buttons (square box look); report list: two-line label
st.markdown("""
<style>
    [data-app-menu-row] ~ [data-testid="stHorizontalBlock"] button,
    [data-testid="stSidebar"] [data-testid="stHorizontalBlock"]:first-of-type button {
        font-size: 1.05rem !important;
        padding: 0.7rem !important;
        font-weight: 600 !important;
        min-height: 3rem !important;
        border-radius: 10px !important;
        border: 1px solid rgba(49, 51, 63, 0.2) !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05) !important;
    }
    [data-app-menu-row] ~ [data-testid="stHorizontalBlock"] button {
        white-space: pre-line !important;
    }
</style>
""", unsafe_allow_html=True)

# Display title and subtitle per page (Reports list vs viewing a report vs Home)
_on_reports_list = st.session_state.get("view_mode") == "reports" and not st.session_state.get("reports_view_filename")
if _on_reports_list:
    title_to_show = "📂 Reports"
    subtitle_to_show = "Choose a report to view."
else:
    title_to_show = st.session_state.get('report_title', "⚡ REPORT")
    subtitle_to_show = st.session_state.get('report_subtitle', "Generate electrical station data reports with time intervals")
st.title(title_to_show)
st.markdown(subtitle_to_show)

# When Reports is selected but no report chosen yet: show prompt only (no duplicate header)
if _on_reports_list:
    _url_reports_list()
    st.info("Select a report from the list on the left to view it here.")
    st.stop()

# Main content area (skip input checks when viewing a saved report from Reports list)
_viewing_saved_report = bool(st.session_state.get("reports_view_filename") and st.session_state.get("reports_view_entry"))
if not _viewing_saved_report:
    if instructions_file is None:
        st.info("👈 Please upload an Instructions Excel file in the sidebar to get started.")
        st.stop()

    if not station_name or station_name.strip() == "":
        if 'station_names_cache' in st.session_state and len(st.session_state.station_names_cache) > 0:
            st.warning("⚠️ Please select a Station Name from the dropdown")
        else:
            st.warning("⚠️ Please enter or select a Station Name")
        st.stop()

    if dc_file is None:
        st.info("👈 Please upload a DC Excel file in the sidebar to get started.")
        st.stop()

    if not bd_folder_path or bd_folder_path.strip() == "":
        st.info("👈 Please enter the path to the BD folder in the sidebar.")
        st.stop()

    if not scada_column or scada_column.strip() == "":
        st.error("❌ SCADA Column Name is required. Please enter the SCADA column name.")
        st.stop()

    if not bd_sheet or bd_sheet.strip() == "":
        st.error("❌ BD Sheet Name is required. Please enter the BD sheet name.")
        st.stop()

# When viewing a past report from Menu Reports: load it into session and set display keys (once)
_reports_view_filename = st.session_state.get("reports_view_filename")
_reports_view_entry = st.session_state.get("reports_view_entry")
if _reports_view_filename and _reports_view_entry and st.session_state.get("reports_view_active") != _reports_view_filename:
    report_key = f"output_data_report_{_reports_view_filename}"
    report_path = REPORTS_DIR / _reports_view_filename
    if report_path.exists():
        try:
            df_report = pd.read_excel(report_path, engine="openpyxl")
            st.session_state[report_key] = df_report
            st.session_state["display_output_data_key"] = report_key
            st.session_state["display_station_name"] = _reports_view_entry.get("station", "")
            date_f = _reports_view_entry.get("date_from", "")
            date_t = _reports_view_entry.get("date_to", "")
            st.session_state["display_stats"] = {
                "total_days": 0,
                "total_slots": 0,
                "output_rows": _reports_view_entry.get("row_count", 0),
            }
            if date_f and date_t:
                st.session_state["report_title"] = f"⚡ REPORT FROM {date_f} TO {date_t}"
            elif date_f:
                st.session_state["report_title"] = f"⚡ REPORT FROM {date_f}"
            else:
                st.session_state["report_title"] = "⚡ REPORT"
            st.session_state["reports_view_active"] = _reports_view_filename
        except Exception:
            st.session_state["reports_view_active"] = None
    else:
        st.session_state["reports_view_active"] = None

# Display output data BEFORE processing - prevents Streamlit "stale" blur during batch reruns
if 'display_output_data_key' in st.session_state:
    output_data_key = st.session_state['display_output_data_key']
    station_name_display = st.session_state.get('display_station_name', '')
    
    if output_data_key in st.session_state:
        df_output = st.session_state[output_data_key]
        if df_output is not None and not df_output.empty:
            df_output = df_output.fillna("").replace("None", "")
            # Round Diff (MW) to 2 decimals for display
            if 'Diff (MW)' in df_output.columns:
                df_output = df_output.copy()
                df_output['Diff (MW)'] = df_output['Diff (MW)'].apply(
                    lambda x: round(x, 2) if isinstance(x, (int, float)) and x != "" else x
                )
        
        processing = st.session_state.get('processing_in_progress', False)
        
        if df_output is not None and not df_output.empty:
            # Show progress caption only when processing
            if processing:
                proc_config = st.session_state.get('processing_config', {})
                total_slots = proc_config.get('total_slots', 1)
                progress_val = min(0.95, 0.6 + 0.3 * len(df_output) / max(1, total_slots))
                st.progress(progress_val)
                current_date = proc_config.get('current_date', '')
                if current_date:
                    st.caption(f"⏳ Processing day {current_date} — {len(df_output)} rows so far")
                else:
                    st.caption(f"⏳ Processing... {len(df_output)} rows so far")
            
            # Stats only when complete
            if not processing and 'display_stats' in st.session_state:
                stats = st.session_state['display_stats']
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total Days", stats.get('total_days', 0))
                with col2:
                    st.metric("Total Slots", stats.get('total_slots', 0))
                with col3:
                    st.metric("Output Rows", stats.get('output_rows', 0))
                if st.session_state.get("reports_view_active"):
                    _url_report_file(st.session_state["reports_view_active"])  # Keep URL: ?view=report&file=...
            
            # Create dynamic table title based on station and date range
            title_parts = ["Calculation sheet for BD and non compliance of", station_name_display]
            
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
            
            # Search and Download (same row)
            search_key = f"{output_data_key}_search"
            rows_key = f"{output_data_key}_rows_per_page"
            page_key = f"{output_data_key}_page"
            col_search, col_download = st.columns([3, 1])
            with col_search:
                current_search = st.session_state.get(search_key, "")
                search_term = st.text_input(
                    "🔍 Search",
                    value=current_search,
                    placeholder="Search in all columns...",
                    help="Filter rows by searching across all columns",
                    key=search_key
                )
            with col_download:
                # Spacer to align Download button with Search input (match label height)
                st.markdown(
                    '<div style="font-size: 14px; font-weight: 500; color: rgb(49, 51, 63); margin-bottom: 0.25rem; min-height: 1.25rem;">&nbsp;</div>',
                    unsafe_allow_html=True
                )
                viewing_saved = st.session_state.get("reports_view_active")
                if viewing_saved:
                    report_path = REPORTS_DIR / viewing_saved
                    if report_path.exists():
                        try:
                            with open(report_path, "rb") as f:
                                file_data = f.read()
                            st.download_button(
                                label="📥 Download Output File",
                                data=file_data,
                                file_name=viewing_saved,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True,
                                key="download_button_output"
                            )
                        except Exception:
                            pass
                elif 'last_output_file_data' in st.session_state:
                    file_data = st.session_state['last_output_file_data']
                    download_filename = st.session_state.get('last_output_filename', 'output.xlsx')
                    st.download_button(
                        label="📥 Download Output File",
                        data=file_data,
                        file_name=download_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="download_button_output"
                    )
                elif 'last_output_path' in st.session_state:
                    output_path = Path(st.session_state['last_output_path'])
                    if output_path.exists():
                        try:
                            with open(output_path, "rb") as f:
                                file_data = f.read()
                            download_filename = st.session_state.get('last_output_filename', 'output.xlsx')
                            st.download_button(
                                label="📥 Download Output File",
                                data=file_data,
                                file_name=download_filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True,
                                key="download_button_output"
                            )
                        except Exception:
                            pass
            
            # Apply search filter and compute pagination
            rows_per_page = st.session_state.get(rows_key, 25)
            if search_term:
                mask = df_output.astype(str).apply(
                    lambda x: x.str.contains(search_term, case=False, na=False)
                ).any(axis=1)
                df_filtered = df_output[mask].copy()
            else:
                df_filtered = df_output.copy()
            
            total_rows = len(df_filtered)
            total_pages = (total_rows + rows_per_page - 1) // rows_per_page if total_rows > 0 else 1
            current_page = st.session_state.get(page_key, 1)
            if current_page > total_pages:
                current_page = 1
            page_num = current_page
            
            start_idx = (page_num - 1) * rows_per_page
            end_idx = start_idx + rows_per_page
            df_display = df_filtered.iloc[start_idx:end_idx].copy()
            
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
                    height=_table_height(len(df_display)),
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
                    height=_table_height(len(df_display)),
                    hide_index=True
                )
            
            # Bottom row: caption left, Rows per page + Page nav right
            col_caption, col_nav = st.columns([2, 1])
            with col_caption:
                if total_pages > 1:
                    st.caption(f"Showing rows {start_idx + 1} to {min(end_idx, total_rows)} of {total_rows} (Page {page_num}/{total_pages})")
                else:
                    st.caption(f"Showing all {total_rows} rows")
            with col_nav:
                rows_options = [10, 25, 50, 100, 500]
                default_rows = st.session_state.get(rows_key, 25)
                default_idx = rows_options.index(default_rows) if default_rows in rows_options else 1
                col_rows, col_prev, col_info, col_next = st.columns([1.2, 0.6, 0.8, 0.6])
                with col_rows:
                    st.selectbox(
                        "Rows per page",
                        options=rows_options,
                        index=default_idx,
                        help="Number of rows to display per page",
                        key=rows_key,
                        label_visibility="visible"
                    )
                if total_pages > 1:
                    label_style = 'font-size: 14px; font-weight: 500; color: rgb(49, 51, 63); margin-bottom: 0.25rem; min-height: 1.25rem;'
                    with col_prev:
                        st.markdown(f'<div style="{label_style}">Page</div>', unsafe_allow_html=True)
                        prev_clicked = st.button("‹", key=f"{page_key}_prev", help="Previous page", use_container_width=True)
                        if prev_clicked:
                            st.session_state[page_key] = max(1, current_page - 1)
                            st.rerun()
                    with col_info:
                        st.markdown(f'<div style="{label_style}">&nbsp;</div>', unsafe_allow_html=True)
                        st.markdown(
                            f"<div style='display: flex; align-items: center; justify-content: center; min-height: 38px; font-weight: 500; font-size: 14px;'>{current_page}/{total_pages}</div>",
                            unsafe_allow_html=True
                        )
                    with col_next:
                        st.markdown(f'<div style="{label_style}">&nbsp;</div>', unsafe_allow_html=True)
                        next_clicked = st.button("›", key=f"{page_key}_next", help="Next page", use_container_width=True)
                        if next_clicked:
                            st.session_state[page_key] = min(total_pages, current_page + 1)
                            st.rerun()

# Trigger continue processing on next run
if st.session_state.get('processing_in_progress'):
    st.session_state['_run_continue_processing'] = True

# Generate button only on Home (not when viewing a saved report from Reports)
_viewing_report = bool(st.session_state.get("reports_view_active"))
run_generate = False
if not _viewing_report:
    run_generate = st.button("🚀 Generate", type="primary", use_container_width=True) or st.session_state.pop('_run_continue_processing', False)
if run_generate:
    # No bottom progress bar or status text (progress shown in report area only)
    class _DummyProgress:
        def progress(self, _): pass
    class _DummyStatus:
        def text(self, _): pass
    progress_bar = _DummyProgress()
    status_text = _DummyStatus()
    
    try:
        # Use persistent temp dir for incremental updates (survives st.rerun)
        processing_continue = st.session_state.get('processing_in_progress', False)
        if processing_continue:
            temp_path = Path(st.session_state.get('processing_temp_dir', ''))
            if not temp_path.exists():
                st.error("❌ Processing temp dir not found. Please run Generate again.")
                st.session_state.pop('processing_in_progress', None)
                st.stop()
            instructions_path = temp_path / st.session_state.get('processing_instructions_name', instructions_file.name)
            dc_path = temp_path / st.session_state.get('processing_dc_name', dc_file.name if dc_file else '') if st.session_state.get('processing_dc_name') else None
            if dc_path and not dc_path.exists():
                dc_path = None
        else:
            # First run: create persistent temp dir
            temp_base = Path(tempfile.gettempdir()) / "electrical_app"
            temp_base.mkdir(parents=True, exist_ok=True)
            run_id = str(uuid.uuid4())[:8]
            temp_path = temp_base / run_id
            temp_path.mkdir(exist_ok=True)
            
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
        
        # Handle BD folder (runs for both first run and continue)
        bd_folder = None
        if bd_folder_path:
            bd_folder = Path(bd_folder_path)
            if not bd_folder.exists():
                if instructions_file.name:
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
        
        # Load or init output_rows for incremental display
        if processing_continue:
            output_rows = st.session_state.get('processing_output_rows', [])
            dc_found_count = st.session_state.get('processing_dc_found', 0)
            dc_not_found_count = st.session_state.get('processing_dc_not_found', 0)
            scada_found_count = st.session_state.get('processing_scada_found', 0)
            scada_not_found_count = st.session_state.get('processing_scada_not_found', 0)
            slots_to_skip = len(output_rows)
        else:
            output_rows = []
            dc_found_count = dc_not_found_count = scada_found_count = scada_not_found_count = 0
            slots_to_skip = 0
        
        total_slots = 0
        processed_slots = 0
        current_date = None
        previous_date_with_data = None
        date_start_row = None
        row_idx = start_data_row
        
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
        
        # Process matches - accumulate to output_rows, batch and rerun for incremental display
        batch_count = 0
        for idx, (row_num, row_data) in enumerate(matches, 1):
            if from_time_col and to_time_col and from_time_col <= len(row_data) and to_time_col <= len(row_data):
                from_time_val = row_data[from_time_col - 1] if from_time_col > 0 else None
                to_time_val = row_data[to_time_col - 1] if to_time_col > 0 else None
                date_val = row_data[date_col - 1] if date_col and date_col > 0 and date_col <= len(row_data) else None
                
                if from_time_val is not None and to_time_val is not None:
                    slots = slots_15min(from_time_val, to_time_val)
                    if slots:
                        date_str = format_value(date_val) if date_val else ""
                        
                        if date_str and date_str != previous_date_with_data and previous_date_with_data is not None:
                            date_start_row = None
                        
                        if date_str and date_str != current_date:
                            current_date = date_str
                            previous_date_with_data = date_str
                            date_start_row = row_idx
                        
                        for slot_idx, (slot_from, slot_to) in enumerate(slots):
                            # Skip already-processed slots when resuming
                            if processed_slots < slots_to_skip:
                                processed_slots += 1
                                row_idx += 1
                                continue
                            
                            row_date = date_str if (slot_idx == 0 and date_str and date_start_row == row_idx) else ""
                            
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
                            
                            # SCADA lookup
                            scada_value = None
                            if scada_cache and date_str:
                                scada_value = find_scada_value(scada_cache, date_str, slot_from, debug=verbose, show_progress=False)
                                if scada_value is not None:
                                    scada_found_count += 1
                                else:
                                    scada_not_found_count += 1
                            
                            # Calculate difference (rounded to 2 decimals)
                            diff_value = None
                            if dc_value is not None and scada_value is not None:
                                try:
                                    dc_num = float(dc_value) if isinstance(dc_value, (int, float, str)) and str(dc_value).strip() else None
                                    scada_num = float(scada_value) if isinstance(scada_value, (int, float, str)) and str(scada_value).strip() else None
                                    if dc_num is not None and scada_num is not None:
                                        diff_value = round(dc_num - scada_num, 2)
                                except (ValueError, TypeError):
                                    pass
                            
                            output_rows.append({
                                'Date': row_date,
                                'From': slot_from,
                                'To': slot_to,
                                'DC (MW)': dc_value if dc_value is not None else "",
                                'As per SLDC Scada in MW': scada_value if scada_value is not None else "",
                                'Diff (MW)': diff_value if diff_value is not None else ""
                            })
                            
                            row_idx += 1
                            processed_slots += 1
                            batch_count += 1
                            
                            # Update progress
                            if total_slots > 0:
                                progress = 60 + int(30 * processed_slots / total_slots)
                                progress_bar.progress(min(progress, 90))
                            
                            # Batch checkpoint: save and rerun for incremental display
                            if batch_count >= PROCESSING_BATCH_SIZE and processed_slots < total_slots:
                                st.session_state['processing_in_progress'] = True
                                st.session_state['processing_output_rows'] = output_rows
                                st.session_state['processing_temp_dir'] = str(temp_path)
                                st.session_state['processing_instructions_name'] = instructions_path.name
                                st.session_state['processing_dc_name'] = dc_path.name if dc_path else None
                                st.session_state['processing_config'] = {
                                    'total_slots': total_slots,
                                    'station_name': station_name,
                                    'current_date': date_str or '',
                                }
                                # Use same display as final - table updates in place
                                partial_key = 'output_data_processing'
                                st.session_state[partial_key] = pd.DataFrame(output_rows).fillna("").replace("None", "")
                                st.session_state['display_output_data_key'] = partial_key
                                st.session_state['display_station_name'] = station_name
                                st.session_state['processing_dc_found'] = dc_found_count
                                st.session_state['processing_dc_not_found'] = dc_not_found_count
                                st.session_state['processing_scada_found'] = scada_found_count
                                st.session_state['processing_scada_not_found'] = scada_not_found_count
                                wb.close()
                                if dc_wb:
                                    dc_wb.close()
                                if scada_cache:
                                    scada_cache.close_all()
                                st.rerun()
        
        # Done - build Excel from output_rows
        wb.close()
        if dc_wb:
            dc_wb.close()
        if scada_cache:
            scada_cache.close_all()
        
        # Clear processing state
        st.session_state.pop('processing_in_progress', None)
        st.session_state.pop('processing_output_rows', None)
        st.session_state.pop('processing_config', None)
        st.session_state.pop('output_data_processing', None)
        
        # Build output Excel from output_rows
        output_wb = openpyxl.Workbook()
        output_sheet = output_wb.active
        output_sheet.title = "Time Intervals"
        headers = ['Date', 'From', 'To', 'DC (MW)', 'As per SLDC Scada in MW', 'Diff (MW)']
        for c, h in enumerate(headers):
            cell = output_sheet.cell(row=header_row_out, column=start_col + c)
            cell.value = h
            cell.font = header_font
            cell.alignment = center_align
            cell.border = thin_border
        
        date_start_row = None
        row_idx = start_data_row
        for r, row_dict in enumerate(output_rows):
            out_row = row_idx + r
            if row_dict['Date'] and date_start_row is not None and out_row > date_start_row:
                output_sheet.merge_cells(f"{get_column_letter(start_col)}{date_start_row}:{get_column_letter(start_col)}{out_row - 1}")
            if row_dict['Date']:
                date_start_row = out_row
            output_sheet.cell(row=out_row, column=start_col).value = row_dict['Date'] or ""
            output_sheet.cell(row=out_row, column=start_col + 1).value = row_dict['From']
            output_sheet.cell(row=out_row, column=start_col + 2).value = row_dict['To']
            output_sheet.cell(row=out_row, column=start_col + 3).value = row_dict['DC (MW)']
            output_sheet.cell(row=out_row, column=start_col + 4).value = row_dict['As per SLDC Scada in MW']
            output_sheet.cell(row=out_row, column=start_col + 5).value = row_dict['Diff (MW)']
            for c in range(6):
                output_sheet.cell(row=out_row, column=start_col + c).border = thin_border
        if date_start_row is not None:
            last_data_row = row_idx + len(output_rows) - 1
            if last_data_row > date_start_row:
                output_sheet.merge_cells(f"{get_column_letter(start_col)}{date_start_row}:{get_column_letter(start_col)}{last_data_row}")
        
        last_row = row_idx + len(output_rows) - 1
        
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
        
        # Store output filename and path in session_state for persistence across reruns
        st.session_state['last_output_filename'] = output_filename
        st.session_state['last_output_path'] = str(output_path)
        # Also store file data in session_state so it persists even if temp file is deleted
        with open(output_path, "rb") as f:
            st.session_state['last_output_file_data'] = f.read()
        
        # Persist to Menu Reports: copy to reports dir and append to index
        try:
            _reports_save_file(Path(output_path), output_filename)
            report_title = st.session_state.get('report_title', '')
            date_from = date_to = ""
            if "FROM" in report_title:
                part = report_title.split("FROM", 1)[1].strip()
                if " TO " in part:
                    date_from, date_to = (s.strip() for s in part.split(" TO ", 1))
                else:
                    date_from = part
            _reports_append_entry({
                "filename": output_filename,
                "station": station_name,
                "date_from": date_from,
                "date_to": date_to,
                "run_at": datetime.now().isoformat(),
                "row_count": len(output_rows),
            })
        except Exception:
            pass  # Don't fail the run if reports persist fails
        
        progress_bar.progress(100)
        status_text.text("✅ Processing complete!")
        
        # Close workbooks
        wb.close()
        if dc_wb:
            dc_wb.close()
        if scada_cache:
            scada_cache.close_all()
        
        # Display summary
        st.success("✅ Output file generated successfully!")
        
        # Load output data for display - store in session state to persist across reruns
        output_data_key = f"output_data_{output_filename}"
        
        try:
            df_output = pd.read_excel(output_path, engine='openpyxl')
            st.session_state[output_data_key] = df_output
            st.session_state[f"{output_data_key}_path"] = str(output_path)
        except Exception as e:
            st.error(f"Could not load output data: {e}")
            df_output = None
        
        # Store reference and stats for display outside this block
        if df_output is not None and not df_output.empty:
            st.session_state['display_output_data_key'] = output_data_key
            st.session_state['display_station_name'] = station_name
            # Clear Menu Reports view so this run becomes the main display
            for key in ("reports_view_filename", "reports_view_entry", "reports_view_active", "reports_view_from_list", "view_mode"):
                st.session_state.pop(key, None)
            _url_main()  # Clean URL for main page
            # Store stats for persistent display
            total_days = len({r['Date'] for r in output_rows if r.get('Date')})
            display_stats = {
                'total_days': total_days,
                'total_slots': total_slots,
                'output_rows': len(output_rows),
                'dc_found': dc_found_count if dc_wb else None,
                'dc_not_found': dc_not_found_count if dc_wb else None,
                'scada_found': scada_found_count if scada_cache else None,
                'scada_not_found': scada_not_found_count if scada_cache else None,
            }
            st.session_state['display_stats'] = display_stats
            # Persist "latest" so "Back to latest" can restore
            st.session_state['last_display_station_name'] = station_name
            st.session_state['last_display_stats'] = display_stats
            st.session_state['last_report_title'] = st.session_state.get('report_title', '⚡ REPORT')
            # Rerun so display block refreshes without processing caption (processing_in_progress was cleared)
            st.rerun()
    
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        if verbose:
            import traceback
            st.code(traceback.format_exc())

