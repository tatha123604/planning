"""
Data loading and transformation helpers for the CLI Matrix updater app.
"""
from __future__ import annotations

import io
from copy import copy
from typing import Dict, Iterable, Tuple, Union

import openpyxl
import pandas as pd

BytesLike = Union[bytes, bytearray, io.BytesIO]


def _as_bytesio(obj: Union[str, BytesLike]) -> io.BytesIO:
    """
    Return a fresh BytesIO for paths, bytes, or file-like uploads.
    A fresh buffer is used each time so the caller can read repeatedly.
    """
    if isinstance(obj, (bytes, bytearray)):
        return io.BytesIO(obj)
    if isinstance(obj, io.BytesIO):
        return io.BytesIO(obj.getvalue())
    if isinstance(obj, str):
        with open(obj, "rb") as fh:
            return io.BytesIO(fh.read())
    # Streamlit file_uploader passes an UploadedFile with .read()
    if hasattr(obj, "read"):
        data = obj.read()
        return io.BytesIO(data)
    raise TypeError("Unsupported input type for file-like data")


def _load_source_df(source_file: Union[str, BytesLike]) -> pd.DataFrame:
    """
    Load the new CLI matrix (source) and normalize column names / types.
    """
    stream = _as_bytesio(source_file)
    df = pd.read_excel(stream, skiprows=2)
    df.columns = [str(c).strip().replace("\n", " ") for c in df.columns]
    df = df.rename(
        columns={
            "AllotedDesig.": "AllotedDesig",
            "Oldest FPOverDue Date": "OldestFP",
            "Oldest CounselOverDue Date": "OldestCounsel",
            "Oldest GradingOverDue": "OldestGrading",
            "TotalOverDue Cases": "TotalOverDue",
        }
    )

    numeric_cols = ["FPOverDue", "CounselOverDue", "GradingOverDue", "TotalOverDue"]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    date_cols = ["OldestFP", "OldestCounsel", "OldestGrading"]
    for col in date_cols:
        df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True).dt.date

    df = df[df["CLI ID"].notna()]
    df["TotalCalc"] = df[["FPOverDue", "CounselOverDue", "GradingOverDue"]].sum(axis=1)
    return df


def _aggregate(df: pd.DataFrame) -> pd.DataFrame:
    """
    Group by CLI + name + designation and aggregate overdue metrics.
    """
    agg = (
        df.groupby(["CLI ID", "CLI Name", "AllotedDesig"], dropna=False)
        .agg(
            FPOverDue=("FPOverDue", "sum"),
            OldestFP=("OldestFP", "min"),
            CounselOverDue=("CounselOverDue", "sum"),
            OldestCounsel=("OldestCounsel", "min"),
            GradingOverDue=("GradingOverDue", "sum"),
            OldestGrading=("OldestGrading", "min"),
            TotalCalc=("TotalCalc", "sum"),
        )
        .reset_index()
    )
    return agg


def build_summary_df(agg: pd.DataFrame) -> pd.DataFrame:
    """
    Sheet 1: only FP overdue rows.
    """
    summary = agg[agg["FPOverDue"] > 0].copy()
    summary = summary[
        ["CLI ID", "CLI Name", "AllotedDesig", "FPOverDue", "OldestFP"]
    ]
    summary.insert(0, "S.No.", range(1, len(summary) + 1))
    summary = summary.rename(
        columns={
            "CLI ID": "CLI ID",
            "CLI Name": "CLI Name",
            "AllotedDesig": "Alloted Desig.",
            "FPOverDue": "FP Over Due",
            "OldestFP": "Oldest FP OverDue Date",
        }
    )
    return summary


def build_master_df(agg: pd.DataFrame) -> pd.DataFrame:
    """
    Sheet 2: all overdue metrics per CLI + designation.
    """
    master = agg.copy()
    master.insert(0, "S.No.", range(1, len(master) + 1))
    master = master.rename(
        columns={
            "CLI ID": "CLI ID",
            "CLI Name": "CLI Name",
            "AllotedDesig": "Alloted Desig.",
            "FPOverDue": "FP Over Due",
            "OldestFP": "Oldest FP OverDue Date",
            "CounselOverDue": "Counsel Over Due",
            "OldestCounsel": "Oldest Counsel OverDue Date",
            "GradingOverDue": "Grading OverDue",
            "OldestGrading": "Oldest Grading OverDue",
            "TotalCalc": "Total Over Due Cases",
        }
    )
    return master


def _read_other_sheets(
    template_file: Union[str, BytesLike], names: Iterable[str]
) -> Dict[str, pd.DataFrame]:
    """
    Read sheets we are not regenerating. header=None to keep raw grid.
    """
    buf = _as_bytesio(template_file)
    xls = pd.ExcelFile(buf)
    return {name: pd.read_excel(xls, sheet_name=name, header=None) for name in names}


def build_output_workbook(
    source_file: Union[str, BytesLike],
    template_file: Union[str, BytesLike],
) -> Tuple[io.BytesIO, pd.DataFrame, pd.DataFrame]:
    """
    Create updated workbook bytes plus the two generated DataFrames.
    """
    src_df = _load_source_df(source_file)
    agg = _aggregate(src_df)
    summary_df = build_summary_df(agg)
    master_df = build_master_df(agg)

    template_buf = _as_bytesio(template_file)
    wb = openpyxl.load_workbook(template_buf)
    summary_sheet = wb.sheetnames[0]
    master_sheet = wb.sheetnames[1]

    _write_summary_sheet(wb[summary_sheet], summary_df)
    _write_master_sheet(wb[master_sheet], master_df)

    # ensure formulas recalc when the user opens the file
    wb.calculation.fullCalcOnLoad = True

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, summary_df, master_df


def _copy_style(src_cell, dest_cell):
    dest_cell._style = copy(src_cell._style)
    dest_cell.font = copy(src_cell.font)
    dest_cell.border = copy(src_cell.border)
    dest_cell.fill = copy(src_cell.fill)
    dest_cell.number_format = src_cell.number_format
    dest_cell.protection = copy(src_cell.protection)
    dest_cell.alignment = copy(src_cell.alignment)


def _clear_data_area(ws, start_row: int, end_row: int, end_col: int):
    for row in range(start_row, end_row + 1):
        for col in range(1, end_col + 1):
            ws.cell(row=row, column=col).value = None


def _apply_row_style(ws, template_row: int, target_row: int, end_col: int):
    for col in range(1, end_col + 1):
        _copy_style(ws.cell(template_row, col), ws.cell(target_row, col))


def _merge_same_cli(ws, start_row: int, rows_count: int, cli_col: int, name_col: int):
    """
    Merge cells vertically for consecutive identical CLI IDs (and names).
    """
    if rows_count == 0:
        return
    data_end = start_row + rows_count - 1
    current_start = start_row
    last_value = ws.cell(start_row, cli_col).value
    for r in range(start_row + 1, data_end + 1):
        val = ws.cell(r, cli_col).value
        if val != last_value:
            if last_value is not None and r - 1 > current_start:
                ws.merge_cells(start_row=current_start, end_row=r - 1, start_column=cli_col, end_column=cli_col)
                ws.merge_cells(start_row=current_start, end_row=r - 1, start_column=name_col, end_column=name_col)
            current_start = r
            last_value = val
    # tail
    if last_value is not None and data_end > current_start:
        ws.merge_cells(start_row=current_start, end_row=data_end, start_column=cli_col, end_column=cli_col)
        ws.merge_cells(start_row=current_start, end_row=data_end, start_column=name_col, end_column=name_col)


def _write_summary_sheet(ws, df: pd.DataFrame):
    """
    Fill Sheet 1 preserving template styling.
    Total row is placed immediately after the data to avoid overwriting entries.
    """
    header_row = 2
    data_start = 3
    cols = ["S.No.", "CLI ID", "CLI Name", "Alloted Desig.", "FP Over Due", "Oldest FP OverDue Date"]
    end_col = len(cols)

    # locate template total row (e.g., row 22) to copy style from
    template_total_row = None
    for r in range(data_start, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if isinstance(val, str) and "total" in val.lower():
            template_total_row = r
            break
    if template_total_row is None:
        template_total_row = data_start + 19  # default fallback

    # clear existing merges except header/title
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row >= data_start:
            ws.unmerge_cells(range_string=str(rng))

    # clear old data area (data + previous total)
    clear_end = max(ws.max_row, data_start + len(df) + 2)
    _clear_data_area(ws, data_start, clear_end, end_col)

    # ensure enough styled rows by copying style from first data row (3)
    style_source_row = data_start
    for idx in range(df.shape[0]):
        target_row = data_start + idx
        _apply_row_style(ws, style_source_row, target_row, end_col)

    # write data
    for i in range(df.shape[0]):
        target_row = data_start + i
        for j, col in enumerate(cols):
            ws.cell(row=target_row, column=j + 1).value = df.iloc[i][col]

    # merge repeated CLI IDs / names
    _merge_same_cli(ws, data_start, df.shape[0], cli_col=2, name_col=3)

    # total row just after data
    total_row = data_start + df.shape[0]
    # copy style from template total row
    _apply_row_style(ws, template_total_row, total_row, end_col)
    # total row styling and value
    ws.cell(total_row, 1).value = "TOTAL FP DUE"
    # use Excel formula to avoid any rounding mismatch and stay accurate if users edit values
    ws.cell(total_row, 5).value = f"=SUM(E{data_start}:E{total_row-1})"
    ws.cell(total_row, 2).value = None
    ws.cell(total_row, 3).value = None
    ws.cell(total_row, 4).value = None
    ws.cell(total_row, 6).value = None

    # re-apply merges for total row
    ws.merge_cells(start_row=total_row, end_row=total_row, start_column=1, end_column=4)
    ws.merge_cells(start_row=total_row, end_row=total_row, start_column=5, end_column=6)


def _write_master_sheet(ws, df: pd.DataFrame):
    """
    Fill Sheet 2 preserving template styling and merge repeated CLI IDs / names.
    """
    header_row = 2
    data_start = 3
    cols = [
        "S.No.",
        "CLI ID",
        "CLI Name",
        "Alloted Desig.",
        "FP Over Due",
        "Oldest FP OverDue Date",
        "Counsel Over Due",
        "Oldest Counsel OverDue Date",
        "Grading OverDue",
        "Oldest Grading OverDue",
        "Total Over Due Cases",
    ]
    end_col = len(cols)

    # clear merges except title row
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row >= data_start:
            ws.unmerge_cells(range_string=str(rng))

    _clear_data_area(ws, data_start, max(ws.max_row, data_start + df.shape[0]), end_col)

    style_source_row = data_start
    for idx in range(df.shape[0]):
        target_row = data_start + idx
        _apply_row_style(ws, style_source_row, target_row, end_col)

    for i in range(df.shape[0]):
        target_row = data_start + i
        for j, col in enumerate(cols):
            ws.cell(row=target_row, column=j + 1).value = df.iloc[i][col]

    _merge_same_cli(ws, data_start, df.shape[0], cli_col=2, name_col=3)
