"""Gera um Excel no molde do arquivo antigo, com a TIMETABLE já atualizada."""

from __future__ import annotations

from io import BytesIO
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from compare import CompareResult, _cell_str, detect_header_row

FILL_CHANGED = PatternFill(fill_type="solid", fgColor="FFF2CC")
FILL_ADDED = PatternFill(fill_type="solid", fgColor="C6EFCE")


def _header_row(ws: Worksheet) -> int:
    max_col = ws.max_column or 1
    max_scan = min(ws.max_row or 1, 20)
    rows = []
    for r in range(1, max_scan + 1):
        rows.append([ws.cell(r, c).value for c in range(1, max_col + 1)])
    if not rows:
        return 1
    raw = pd.DataFrame(rows)
    return detect_header_row(raw) + 1


def _headers(ws: Worksheet, header_row: int | None = None) -> dict[str, int]:
    row = header_row or _header_row(ws)
    mapping: dict[str, int] = {}
    for col in range(1, (ws.max_column or 1) + 1):
        value = ws.cell(row, col).value
        if value is None or str(value).strip() == "":
            continue
        mapping[str(value).strip()] = col
    return mapping


def _coerce_value(text: str, sample: Any = None) -> Any:
    if text == "":
        return None
    if hasattr(sample, "strftime"):
        try:
            import pandas as pd

            ts = pd.to_datetime(text, errors="coerce")
            if pd.notna(ts):
                return ts.to_pydatetime()
        except Exception:
            pass
    if text.lstrip("-").isdigit():
        try:
            return int(text)
        except ValueError:
            return text
    try:
        if "." in text:
            return float(text)
    except ValueError:
        pass
    return text


def _write_cell(ws: Worksheet, row: int, col: int, text: str, fill: PatternFill | None) -> None:
    cell = ws.cell(row, col)
    cell.value = _coerce_value(text, cell.value)
    if fill is not None:
        cell.fill = fill


def _copy_header_style(ws: Worksheet, from_col: int, to_col: int, header_row: int = 1) -> None:
    src = ws.cell(header_row, from_col)
    dst = ws.cell(header_row, to_col)
    if src.has_style:
        dst.font = src.font.copy()
        dst.alignment = src.alignment.copy()
        dst.border = src.border.copy()
        dst.fill = src.fill.copy()
    letter = get_column_letter(from_col)
    if ws.column_dimensions[letter].width:
        ws.column_dimensions[get_column_letter(to_col)].width = ws.column_dimensions[letter].width


def _ensure_column(
    ws: Worksheet, headers: dict[str, int], name: str, header_row: int = 1
) -> int:
    if name in headers:
        return headers[name]
    col = (ws.max_column or 0) + 1
    if headers:
        _copy_header_style(ws, next(iter(headers.values())), col, header_row)
    ws.cell(header_row, col).value = name
    headers[name] = col
    return col


def _last_data_row(ws: Worksheet) -> int:
    max_r = ws.max_row or 1
    for row in range(max_r, 1, -1):
        if any(ws.cell(row, col).value not in (None, "") for col in range(1, (ws.max_column or 1) + 1)):
            return row
    return 1


def _row_matches(ws: Worksheet, excel_row: int, headers: dict[str, int], values: dict[str, str]) -> bool:
    for col_name, expected in values.items():
        col = headers.get(col_name)
        if col is None:
            return False
        if _cell_str(ws.cell(excel_row, col).value) != expected:
            return False
    return True


def _find_rows(
    ws: Worksheet, headers: dict[str, int], match: dict[str, str], header_row: int = 1
) -> list[int]:
    found: list[int] = []
    for row in range(header_row + 1, _last_data_row(ws) + 1):
        if _row_matches(ws, row, headers, match):
            found.append(row)
    return found


def _load_template(old_bytes: bytes, sheet_name: str) -> Workbook:
    try:
        wb = load_workbook(BytesIO(old_bytes))
    except Exception:
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
        return wb
    if sheet_name not in wb.sheetnames:
        wb.create_sheet(sheet_name)
    return wb


def build_updated_excel(
    old_bytes: bytes,
    result: CompareResult,
    sheet_name: str,
    key_cols: list[str] | None = None,
) -> bytes:
    """Copia o arquivo antigo e aplica as diferenças na aba informada."""
    wb = _load_template(old_bytes, sheet_name)
    ws = wb[sheet_name]
    header_row = _header_row(ws)
    headers = _headers(ws, header_row)

    for change in result.modified:
        col = _ensure_column(ws, headers, change.column, header_row)
        _write_cell(ws, change.row, col, change.new, FILL_CHANGED)

    added = result.added_rows.copy()
    if "_linha_excel" in added.columns:
        added = added.drop(columns=["_linha_excel"])

    next_row = _last_data_row(ws) + 1
    for _, row in added.iterrows():
        for col_name, value in row.items():
            col = _ensure_column(ws, headers, str(col_name), header_row)
            _write_cell(ws, next_row, col, _cell_str(value), FILL_ADDED)
        next_row += 1

    removed = result.removed_rows.copy()
    rows_to_delete: list[int] = []
    if "_linha_excel" in removed.columns:
        rows_to_delete = sorted({int(v) for v in removed["_linha_excel"]}, reverse=True)
        removed = removed.drop(columns=["_linha_excel"])
    elif key_cols:
        for _, row in removed.iterrows():
            match = {k: _cell_str(row[k]) for k in key_cols if k in row.index}
            rows_to_delete.extend(_find_rows(ws, headers, match, header_row))
        rows_to_delete = sorted(set(rows_to_delete), reverse=True)
    else:
        for _, row in removed.iterrows():
            match = {str(c): _cell_str(row[c]) for c in removed.columns}
            rows_to_delete.extend(_find_rows(ws, headers, match, header_row))
        rows_to_delete = sorted(set(rows_to_delete), reverse=True)

    for excel_row in rows_to_delete:
        ws.delete_rows(excel_row)

    out = BytesIO()
    wb.save(out)
    return out.getvalue()
