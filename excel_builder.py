"""Build report Excel workbook from output rows."""

from typing import Optional

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HEADERS = ["Date", "From", "To", "DC (MW)", "As per SLDC Scada in MW", "Diff (MW)"]
COLUMN_WIDTHS = [15, 10, 10, 12, 25, 12]
PAD = 0


def _make_styles():
    return {
        "header_font": Font(bold=True, size=11),
        "center_align": Alignment(horizontal="center", vertical="center"),
        "thin_border": Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        ),
    }


def build_report_workbook(output_rows: list[dict]) -> Workbook:
    """
    Build and return an openpyxl Workbook with 'Time Intervals' sheet
    filled with output_rows. Caller should save to path.
    """
    wb = Workbook()
    sheet = wb.active
    sheet.title = "Time Intervals"
    styles = _make_styles()
    header_font = styles["header_font"]
    center_align = styles["center_align"]
    thin_border = styles["thin_border"]

    header_row = 1 + PAD
    start_col = 1 + PAD
    start_data_row = 2 + PAD

    # Header row
    for c, h in enumerate(HEADERS):
        cell = sheet.cell(row=header_row, column=start_col + c)
        cell.value = h
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border

    # Data rows and date merging
    date_start_row: Optional[int] = None
    row_idx = start_data_row
    for r, row_dict in enumerate(output_rows):
        out_row = row_idx + r
        if row_dict.get("Date") and date_start_row is not None and out_row > date_start_row:
            sheet.merge_cells(
                f"{get_column_letter(start_col)}{date_start_row}:{get_column_letter(start_col)}{out_row - 1}"
            )
        if row_dict.get("Date"):
            date_start_row = out_row
        sheet.cell(row=out_row, column=start_col).value = row_dict.get("Date") or ""
        sheet.cell(row=out_row, column=start_col + 1).value = row_dict.get("From")
        sheet.cell(row=out_row, column=start_col + 2).value = row_dict.get("To")
        sheet.cell(row=out_row, column=start_col + 3).value = row_dict.get("DC (MW)")
        sheet.cell(row=out_row, column=start_col + 4).value = row_dict.get("As per SLDC Scada in MW")
        sheet.cell(row=out_row, column=start_col + 5).value = row_dict.get("Diff (MW)")
        for c in range(6):
            sheet.cell(row=out_row, column=start_col + c).border = thin_border

    if date_start_row is not None:
        last_data_row = row_idx + len(output_rows) - 1
        if last_data_row > date_start_row:
            sheet.merge_cells(
                f"{get_column_letter(start_col)}{date_start_row}:{get_column_letter(start_col)}{last_data_row}"
            )

    last_row = row_idx + len(output_rows) - 1
    last_content_col = start_col + 5

    sheet.freeze_panes = sheet.cell(row=start_data_row, column=start_col).coordinate
    sheet.sheet_view.showGridLines = False

    for i, w in enumerate(COLUMN_WIDTHS):
        sheet.column_dimensions[get_column_letter(start_col + i)].width = w

    sheet.print_area = f"A1:{get_column_letter(last_content_col)}{last_row}"

    return wb
