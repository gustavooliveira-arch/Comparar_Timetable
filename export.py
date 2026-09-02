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
    """Copia o arquivo antigo por inteiro e aplica as diferenças na aba informada.

    Mantém todas as linhas originais (alteradas, inalteradas, etc.), apenas
    marcando com cor o que mudou/foi adicionado e removendo o que foi excluído.
    """
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


def build_diff_only_excel(
    old_bytes: bytes,
    result: CompareResult,
    sheet_name: str,
) -> bytes:
    """Gera um Excel enxuto: mantém a config da tabela antiga (nomes de
    coluna, estilo e largura do cabeçalho), mas só inclui os dados que
    mudaram.

    Regras:
    - O número total de linhas de dado é o mesmo do arquivo antigo (cada
      linha mantém a mesma posição relativa que tinha na tabela antiga,
      só reindexada para o cabeçalho começar na linha 1). Linhas sem
      nenhuma alteração ficam presentes, porém totalmente em branco.
    - Linhas MODIFICADAS: só as células que de fato mudaram são
      preenchidas (com destaque amarelo); as demais colunas dessa linha
      ficam em branco.
    - Linhas ADICIONADAS: não existiam antes, então são anexadas depois
      da última linha original (aumentando o total). Como são novas por
      inteiro, todas as suas colunas são escritas (com destaque verde).
    - Linhas REMOVIDAS: ficam em branco (não são apagadas, para não
      deslocar a posição das demais linhas).
    """
    old_wb = load_workbook(BytesIO(old_bytes))
    if sheet_name not in old_wb.sheetnames:
        raise ValueError(f"Aba '{sheet_name}' não encontrada no arquivo antigo.")
    old_ws = old_wb[sheet_name]
    header_row = _header_row(old_ws)
    headers = _headers(old_ws, header_row)

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    # Copia o cabeçalho (nomes, estilo de fonte/borda/preenchimento e
    # largura de coluna) da aba antiga para a linha 1 da planilha nova.
    max_col = old_ws.max_column or 1
    for col in range(1, max_col + 1):
        src = old_ws.cell(header_row, col)
        dst = ws.cell(1, col)
        dst.value = src.value
        if src.has_style:
            dst.font = src.font.copy()
            dst.alignment = src.alignment.copy()
            dst.border = src.border.copy()
            dst.fill = src.fill.copy()
        letter = get_column_letter(col)
        if old_ws.column_dimensions[letter].width:
            ws.column_dimensions[letter].width = old_ws.column_dimensions[letter].width
    if old_ws.row_dimensions[header_row].height:
        ws.row_dimensions[1].height = old_ws.row_dimensions[header_row].height

    headers_new: dict[str, int] = dict(headers)
    next_free_col = max_col + 1

    def ensure_col(name: str) -> int:
        nonlocal next_free_col
        if name in headers_new:
            return headers_new[name]
        col = next_free_col
        next_free_col += 1
        ref_col = next(iter(headers_new.values()), 1)
        _copy_header_style(ws, ref_col, col, header_row=1)
        ws.cell(1, col).value = name
        headers_new[name] = col
        return col

    # O cabeçalho do arquivo antigo pode não estar na linha 1 (ex.: tem
    # linhas de título acima). Aqui ele sempre vai para a linha 1 da
    # planilha nova, então todo número de linha de dado precisa ser
    # deslocado pela mesma diferença, preservando a posição relativa de
    # cada linha (e portanto o total de linhas de dados original).
    row_offset = header_row - 1
    last_old_data_row = _last_data_row(old_ws)
    next_out_row = max(last_old_data_row - row_offset, 1)  # onde as linhas adicionadas começam

    for change in result.modified:
        target_row = change.row - row_offset
        col = ensure_col(change.column)
        _write_cell(ws, target_row, col, change.new, FILL_CHANGED)

    # Linhas adicionadas: não existiam no arquivo antigo, então são
    # anexadas depois da última linha de dado original (aumentando o
    # total). Escreve a linha inteira (todas as colunas).
    added = result.added_rows.copy()
    if "_linha_excel" in added.columns:
        added = added.drop(columns=["_linha_excel"])

    for _, row in added.iterrows():
        next_out_row += 1
        for col_name, value in row.items():
            col = ensure_col(str(col_name))
            _write_cell(ws, next_out_row, col, _cell_str(value), FILL_ADDED)

    out = BytesIO()
    wb.save(out)
    return out.getvalue()