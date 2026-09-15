"""Exports for the complete RTIS table, using the project's report styling."""
from datetime import date, datetime
from io import BytesIO
from typing import Iterable, Sequence

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .table_pdf import _build_table_pdf_bytes


def _excel_value(header: str, value: str):
    if not value:
        return None, "General"
    try:
        if header == "Sr.No.":
            return int(value), "0"
        if header in ("Latitude", "Longitude", "Speed"):
            return float(value), "0.######"
        if header in ("Event Time", "Reporting Time"):
            return datetime.fromisoformat(value), "yyyy-mm-dd hh:mm:ss"
        if header == "Train Start Date":
            return date.fromisoformat(value), "yyyy-mm-dd"
    except (ValueError, OverflowError):
        pass
    # Identifiers and unrecognized source values remain literal text.
    return value, "@"


def build_rtis_excel(headers: Sequence[str], rows: Iterable[list[str]]) -> bytes:
    workbook = Workbook(write_only=True)
    sheet = workbook.create_sheet("RTIS Output")
    sheet.freeze_panes = "A2"
    sheet.sheet_view.showGridLines = False
    widths = (10, 14, 14, 14, 14, 12, 23, 12, 10, 15, 23, 32, 18)
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    header_font = Font(name="Arial", size=10, bold=True, color="F3FBFF")
    body_font = Font(name="Arial", size=10, color="16314C")
    header_fill = PatternFill(fill_type="solid", fgColor="16314C")
    header_alignment = Alignment(horizontal="center", vertical="center")
    header_cells = []
    for header in headers:
        cell = WriteOnlyCell(sheet, value=header)
        cell.font, cell.fill, cell.alignment = header_font, header_fill, header_alignment
        header_cells.append(cell)
    sheet.row_dimensions[1].height = 24
    sheet.append(header_cells)
    row_count = 0
    for row in rows:
        cells = []
        for header, raw in zip(headers, row):
            value, number_format = _excel_value(header, raw)
            cell = WriteOnlyCell(sheet, value=value)
            if isinstance(value, str):
                cell.data_type = "s"
            cell.number_format = number_format
            cell.font = body_font
            cells.append(cell)
        sheet.append(cells)
        row_count += 1
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{row_count + 1}"
    stream = BytesIO()
    workbook.save(stream)
    return stream.getvalue()


def build_rtis_pdf(headers: Sequence[str], rows: Iterable[list[str]]) -> bytes:
    return _build_table_pdf_bytes(
        "RTIS Output - All Divisions", list(headers), list(rows),
        "Divisions: SDAH, HWH, ASN, MLDT",
        page_width=1191.0, page_height=842.0, body_font_size=8.5, compress=True,
    )
