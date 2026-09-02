"""Compara a aba TIMETABLE de dois arquivos Excel."""

from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

SHEET_NAME = "TIMETABLE"


def _cell_str(value) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        if value.hour or value.minute or value.second:
            return value.strftime("%Y-%m-%d %H:%M:%S")
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    if text.endswith(".0") and text.replace(".", "", 1).replace("-", "", 1).isdigit():
        try:
            as_float = float(text)
            if as_float.is_integer():
                return str(int(as_float))
        except ValueError:
            pass
    return text


def normalize_frame(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]
    out = out.dropna(how="all")
    for col in out.columns:
        out[col] = out[col].map(_cell_str)
    out = _drop_ghost_columns(out)
    return out.reset_index(drop=True)


def _is_ghost_name(name: str) -> bool:
    text = str(name).strip()
    lower = text.lower()
    return (
        text == ""
        or lower in {"nan", "none", "unnamed"}
        or lower.startswith("unnamed:")
        or lower.startswith("col_")
    )


def _drop_ghost_columns(df: pd.DataFrame) -> pd.DataFrame:
    keep = []
    for col in df.columns:
        empty = df[col].map(_cell_str).eq("").all()
        if empty:
            continue
        keep.append(col)
    return df[keep] if keep else df.iloc[:, 0:0]


def list_sheet_names(path_or_buffer) -> list[str]:
    xl = pd.ExcelFile(path_or_buffer)
    return list(xl.sheet_names)


def resolve_sheet_name(sheet_names: list[str], requested: str) -> str | None:
    wanted = requested.strip().casefold()
    for name in sheet_names:
        if name.strip().casefold() == wanted:
            return name
    return None


def detect_header_row(raw: pd.DataFrame, max_scan: int = 20) -> int:
    """Índice 0-based da linha que parece cabeçalho real (ignora título tipo DISTÂNCIAS)."""
    best_i, best_score = 0, -1.0
    limit = min(max_scan, len(raw))
    for i in range(limit):
        nonempty = [_cell_str(v) for v in raw.iloc[i] if _cell_str(v)]
        if len(nonempty) < 2:
            continue
        unique = len(set(nonempty))
        score = float(unique)
        if unique == len(nonempty):
            score += 2
        if score > best_score:
            best_score = score
            best_i = i
    return best_i


def _unique_headers(names: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    out: list[str] = []
    for name in names:
        text = name.strip() if name else ""
        if _is_ghost_name(text):
            text = "Unnamed"
        n = seen.get(text, 0)
        seen[text] = n + 1
        out.append(text if n == 0 else f"{text}_{n}")
    return out


def _frame_from_raw(raw: pd.DataFrame, header_idx: int) -> pd.DataFrame:
    headers = _unique_headers([_cell_str(v) for v in raw.iloc[header_idx]])
    data = raw.iloc[header_idx + 1 :].copy()
    data.columns = headers
    data = data.dropna(how="all")
    for col in data.columns:
        data[col] = data[col].map(_cell_str)
    data = _drop_ghost_columns(data)
    out = data.reset_index(drop=True)
    out.attrs["header_excel_row"] = header_idx + 1
    return out


def load_timetable(path_or_buffer, sheet_name: str = SHEET_NAME) -> pd.DataFrame:
    xl = pd.ExcelFile(path_or_buffer)
    resolved = resolve_sheet_name(list(xl.sheet_names), sheet_name)
    if resolved is None:
        available = ", ".join(xl.sheet_names)
        raise ValueError(
            f'A aba "{sheet_name}" não existe neste arquivo. '
            f"Abas do Excel (não são colunas): {available}"
        )
    raw = pd.read_excel(xl, sheet_name=resolved, header=None)
    if raw.dropna(how="all").empty:
        df = pd.DataFrame()
        df.attrs["header_excel_row"] = 1
        return df
    raw = raw.reset_index(drop=True)
    return _frame_from_raw(raw, detect_header_row(raw))


@dataclass
class CellChange:
    row: int
    column: str
    old: str
    new: str


@dataclass
class CompareResult:
    added_rows: pd.DataFrame
    removed_rows: pd.DataFrame
    modified: list[CellChange] = field(default_factory=list)
    extra_columns_old: list[str] = field(default_factory=list)
    extra_columns_new: list[str] = field(default_factory=list)
    duplicate_keys_old: int = 0
    duplicate_keys_new: int = 0

    @property
    def has_changes(self) -> bool:
        return bool(
            len(self.added_rows)
            or len(self.removed_rows)
            or self.modified
            or self.extra_columns_old
            or self.extra_columns_new
        )

    def modified_frame(self) -> pd.DataFrame:
        if not self.modified:
            return pd.DataFrame(columns=["linha", "coluna", "antes", "depois"])
        return pd.DataFrame(
            [
                {
                    "linha": c.row,
                    "coluna": c.column,
                    "antes": c.old,
                    "depois": c.new,
                }
                for c in self.modified
            ]
        )


def _header_excel_row(df: pd.DataFrame) -> int:
    return int(df.attrs.get("header_excel_row", 1))


def _extra_columns(left: pd.DataFrame, right: pd.DataFrame) -> list[str]:
    return [
        c
        for c in left.columns
        if c not in right.columns and not _is_ghost_name(str(c))
    ]


def compare_positional(old: pd.DataFrame, new: pd.DataFrame) -> CompareResult:
    """Compara célula a célula, alinhando pela ordem das linhas (mesmo layout)."""
    extra_old = _extra_columns(old, new)
    extra_new = _extra_columns(new, old)
    shared = [c for c in old.columns if c in new.columns]

    max_rows = max(len(old), len(new))
    old_p = old.reindex(range(max_rows)).fillna("")
    new_p = new.reindex(range(max_rows)).fillna("")
    first_data_row = _header_excel_row(old) + 1

    modified: list[CellChange] = []
    for i in range(min(len(old), len(new))):
        for col in shared:
            a, b = old_p.at[i, col], new_p.at[i, col]
            if a != b:
                modified.append(CellChange(row=i + first_data_row, column=col, old=a, new=b))

    added = pd.DataFrame()
    removed = pd.DataFrame()
    if len(new) > len(old):
        added = new.iloc[len(old) :].copy()
        added.insert(0, "_linha_excel", range(first_data_row + len(old), first_data_row + len(new)))
    elif len(old) > len(new):
        removed = old.iloc[len(new) :].copy()
        removed.insert(0, "_linha_excel", range(first_data_row + len(new), first_data_row + len(old)))

    return CompareResult(
        added_rows=added,
        removed_rows=removed,
        modified=modified,
        extra_columns_old=extra_old,
        extra_columns_new=extra_new,
    )


def _row_key(df: pd.DataFrame, keys: list[str]) -> pd.Series:
    if len(keys) == 1:
        return df[keys[0]].astype(str)
    return df[keys].astype(str).agg(" | ".join, axis=1)


def compare_by_keys(old: pd.DataFrame, new: pd.DataFrame, keys: list[str]) -> CompareResult:
    """Compara linhas pela chave (funciona mesmo se a ordem mudou)."""
    missing_old = [k for k in keys if k not in old.columns]
    missing_new = [k for k in keys if k not in new.columns]
    if missing_old or missing_new:
        raise ValueError(
            "Colunas-chave ausentes. "
            f"No arquivo antigo: {missing_old or 'ok'}. "
            f"No arquivo novo: {missing_new or 'ok'}."
        )

    extra_old = _extra_columns(old, new)
    extra_new = _extra_columns(new, old)
    shared = [c for c in old.columns if c in new.columns and c not in keys]

    old_w = old.copy()
    new_w = new.copy()
    old_w["_key"] = _row_key(old_w, keys)
    new_w["_key"] = _row_key(new_w, keys)

    dup_old = int(old_w["_key"].duplicated().sum())
    dup_new = int(new_w["_key"].duplicated().sum())

    old_first = old_w.drop_duplicates("_key", keep="first").set_index("_key")
    new_first = new_w.drop_duplicates("_key", keep="first").set_index("_key")

    old_keys = set(old_first.index)
    new_keys = set(new_first.index)

    removed = old_first.loc[list(old_keys - new_keys)].drop(columns=["_key"], errors="ignore")
    added = new_first.loc[list(new_keys - old_keys)].drop(columns=["_key"], errors="ignore")

    modified: list[CellChange] = []
    for key in sorted(old_keys & new_keys):
        o_row = old_first.loc[key]
        n_row = new_first.loc[key]
        excel_row = int(old_w.index[old_w["_key"] == key][0]) + _header_excel_row(old) + 1
        for col in shared:
            a, b = _cell_str(o_row[col]), _cell_str(n_row[col])
            if a != b:
                modified.append(CellChange(row=excel_row, column=col, old=a, new=b))

    return CompareResult(
        added_rows=added.reset_index(drop=True),
        removed_rows=removed.reset_index(drop=True),
        modified=modified,
        extra_columns_old=extra_old,
        extra_columns_new=extra_new,
        duplicate_keys_old=dup_old,
        duplicate_keys_new=dup_new,
    )