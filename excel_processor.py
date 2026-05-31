"""Excel 读取、筛选与按列拆分。"""
from __future__ import annotations

import os
import re
import tempfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
DATE_KEYWORDS = ("日期", "时间", "年月")
NUMBER_KEYWORDS = ("投资", "金额", "万", "数量", "容量", "长度", "含税", "不含税")
TEXT_ID_KEYWORDS = ("序号", "编号", "文号", "凭证", "WBS", "物料", "人", "名称", "描述")
NON_DATE_KEYWORDS = ("序号", "编号", "年份", "批次", "性质", "类型", "分类")
SUBHEADER_HINTS = ("计划", "实际", "开工", "竣工", "送审", "审定", "关闭", "决算", "容量", "长度")


@dataclass
class ExcelSheet:
    """加载后的工作表：保留原始表头行，数据区用列名索引。"""

    df: pd.DataFrame
    header_rows: list[list[Any]]


@dataclass
class SplitExportResult:
    count: int
    output_path: Path
    mode: str = "files"  # files | workbook


OUTPUT_FOLDER_LABEL = "拆分表格"
OUTPUT_MODE_FILES = "files"
OUTPUT_MODE_WORKBOOK = "workbook"
INVALID_SHEET_CHARS = re.compile(r"[\\/*?:\[\]]")
MAX_SHEET_NAME_LEN = 31


def _cell_str(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _looks_like_data_value(value: Any) -> bool:
    s = _cell_str(value)
    if not s:
        return False
    if len(s) >= 15:
        return True
    digits = s.replace(".", "").replace("-", "")
    if digits.isdigit() and len(digits) >= 8:
        return True
    return False


def _looks_like_subheader(value: Any) -> bool:
    s = _cell_str(value)
    if not s or len(s) > 12:
        return False
    return any(h in s for h in SUBHEADER_HINTS)


def detect_header_row_count(raw: pd.DataFrame) -> int:
    """判断表头为 1 行还是 2 行。"""
    if len(raw) < 2:
        return 1

    row1 = raw.iloc[1]
    row0 = raw.iloc[0]
    r1_nonnull = row1.notna().sum()
    if r1_nonnull == 0:
        return 1

    sample_n = min(12, len(row0))
    data_like = 0
    for i in range(sample_n):
        if pd.notna(row0.iloc[i]) and pd.notna(row1.iloc[i]):
            if _looks_like_data_value(row1.iloc[i]):
                data_like += 1

    if data_like >= 3:
        return 1

    r1_texts = [_cell_str(v) for v in row1 if _cell_str(v)]
    subheader_like = sum(1 for v in row1 if _looks_like_subheader(v))
    r0_nonnull = max(int(row0.notna().sum()), 1)

    if r1_nonnull < r0_nonnull * 0.85 and subheader_like >= 3:
        return 2
    if subheader_like >= max(3, len(r1_texts) * 0.25) and data_like == 0:
        return 2

    return 1


def _dedupe_names(names: list[str]) -> list[str]:
    used: dict[str, int] = {}
    result: list[str] = []
    for name in names:
        base = name or "未命名"
        if base in used:
            used[base] += 1
            result.append(f"{base}_{used[base]}")
        else:
            used[base] = 0
            result.append(base)
    return result


def build_column_labels(header_rows: list[list[Any]]) -> list[str]:
    """生成与 Excel 一致的列名（双行表头优先使用子行名称）。"""
    if len(header_rows) == 1:
        names = [_cell_str(v) or f"列{i + 1}" for i, v in enumerate(header_rows[0])]
        return _dedupe_names(names)

    row0 = pd.Series(header_rows[0]).ffill()
    row1 = header_rows[1]
    names: list[str] = []
    for i in range(len(row0)):
        h1 = row1[i] if i < len(row1) else None
        h0 = row0.iloc[i]
        if _cell_str(h1):
            name = _cell_str(h1)
        elif _cell_str(h0):
            name = _cell_str(h0)
        else:
            name = f"列{i + 1}"
        names.append(name)
    return _dedupe_names(names)


def read_excel_with_headers(path: str | Path, sheet_name: str | int = 0) -> ExcelSheet:
    raw = pd.read_excel(path, header=None, sheet_name=sheet_name)
    if raw.empty:
        raise ValueError("表格为空")

    header_count = detect_header_row_count(raw)
    header_rows = [raw.iloc[r].tolist() for r in range(header_count)]
    data = raw.iloc[header_count:].copy()
    columns = build_column_labels(header_rows)
    data.columns = columns[: len(data.columns)]
    data.reset_index(drop=True, inplace=True)
    return ExcelSheet(df=data, header_rows=header_rows)


def list_sheet_names(path: str | Path) -> list[str]:
    return pd.ExcelFile(path).sheet_names


def _series_looks_like_real_dates(series: pd.Series) -> bool:
    """仅当样本确为真实日期时才判定为日期列（避免序号等小整数误判）。"""
    sample = series.dropna().head(200)
    if sample.empty:
        return False

    if pd.api.types.is_datetime64_any_dtype(series):
        years = pd.to_datetime(sample, errors="coerce").dt.year
        return (years >= 1990).mean() >= 0.6

    as_str = sample.astype(str).str.strip()
    # 工号、编号等纯数字文本（含前导零）不是日期
    if as_str.str.match(r"^\d{5,12}$").mean() >= 0.6:
        return False
    if as_str.str.match(r"^0\d+$").mean() >= 0.3:
        return False
    if as_str.str.match(r"^\d{4}[/\-年]\d{1,2}[/\-月]\d{1,2}").mean() >= 0.6:
        return True
    if as_str.str.match(r"^\d{4}[/\-]\d{1,2}[/\-]\d{1,2}").mean() >= 0.6:
        return True

    numeric = pd.to_numeric(sample, errors="coerce")
    if numeric.notna().mean() >= 0.6:
        # Excel 日期序列号通常 > 10000（约 1982 年以后）
        return (numeric > 10000).mean() >= 0.6

    parsed = pd.to_datetime(sample, errors="coerce", format="mixed")
    valid = parsed.notna()
    if valid.mean() < 0.6:
        return False
    years = parsed[valid].dt.year
    if (years == 1970).mean() > 0.3:
        return False
    return (years >= 1990).mean() >= 0.6


COLUMN_KIND_LABELS = {"date": "日期", "number": "数字", "text": "文本"}
COLUMN_KIND_FROM_LABEL = {v: k for k, v in COLUMN_KIND_LABELS.items()}


def detect_all_column_kinds(df: pd.DataFrame) -> dict[str, str]:
    """自动检测所有列类型，返回 {列名: date|number|text}。"""
    return {str(col): infer_column_kind(df[col], str(col)) for col in df.columns}


def infer_column_kind(series: pd.Series, col_name: str) -> str:
    name = str(col_name)

    if name.endswith("人") or ("制单" in name and "日期" not in name and "时间" not in name):
        return "text"
    if any(k in name for k in TEXT_ID_KEYWORDS):
        return "text"
    if any(k in name for k in NON_DATE_KEYWORDS):
        if "日期" in name or ("时间" in name and "年份" not in name):
            pass
        elif "年份" in name:
            return "number"
        else:
            return "text"

    if "日期" in name:
        return "date"
    if "时间" in name and "年份" not in name:
        return "date"

    if any(k in name for k in NUMBER_KEYWORDS):
        return "number"

    if _series_looks_like_real_dates(series):
        return "date"

    sample = series.dropna().head(200)
    if not sample.empty:
        numeric = pd.to_numeric(sample, errors="coerce")
        if numeric.notna().mean() >= 0.6:
            return "number"

    return "text"


def to_datetime_series(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce", format="mixed")


def to_numeric_series(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def apply_filters(df: pd.DataFrame, filters: list[dict[str, Any]]) -> pd.DataFrame:
    result = df
    for f in filters:
        col = f["column"]
        if col not in result.columns:
            continue
        kind = f["kind"]
        series = result[col]

        if kind == "date" or f.get("start") or f.get("end"):
            start = f.get("start")
            end = f.get("end")
            if not start and not end:
                continue
            dates = to_datetime_series(series)
            mask = dates.notna()
            if start:
                mask &= dates.dt.normalize() >= pd.Timestamp(start).normalize()
            if end:
                mask &= dates.dt.normalize() <= pd.Timestamp(end).normalize()
            result = result[mask]

        elif kind == "number":
            lo = f.get("min")
            hi = f.get("max")
            nums = to_numeric_series(series)
            mask = pd.Series(True, index=result.index)
            if lo is not None and lo != "":
                mask &= nums >= float(lo)
            if hi is not None and hi != "":
                mask &= nums <= float(hi)
            result = result[mask]

        else:
            text = (f.get("text") or "").strip()
            if text:
                result = result[series.astype(str).str.contains(text, na=False, regex=False)]

    return result.reset_index(drop=True)


DATE_EXPORT_FMT = "%Y/%m/%d"
ROW_HEIGHT = 30
FONT_SIZE = 10
MIN_COL_WIDTH = 8
MAX_COL_WIDTH = 80
COL_WIDTH_PADDING = 2


def prepare_dataframe_for_export(
    df: pd.DataFrame,
    col_kinds: dict[str, str] | None = None,
) -> tuple[pd.DataFrame, dict[str, str]]:
    """格式化导出数据，并返回每列类型（date/number/text）。"""
    out = df.copy()
    resolved_kinds: dict[str, str] = {}

    for col in out.columns:
        col_name = str(col)
        if col_kinds and col_name in col_kinds:
            kind = col_kinds[col_name]
        else:
            kind = infer_column_kind(out[col], col_name)
        resolved_kinds[col_name] = kind

        if kind == "date":
            parsed = to_datetime_series(out[col])
            formatted = pd.Series(index=out.index, dtype=object)
            valid = parsed.notna()
            formatted.loc[valid] = parsed.loc[valid].dt.strftime(DATE_EXPORT_FMT)
            formatted.loc[~valid] = None
            out[col] = formatted
        elif kind == "text":
            out[col] = out[col].apply(_to_text_export_value)
        elif kind == "number":
            nums = pd.to_numeric(out[col], errors="coerce")
            if "年份" in str(col):
                out[col] = nums.apply(lambda x: int(x) if pd.notna(x) and x == int(x) else x)
            else:
                out[col] = nums

    return out, resolved_kinds


def _to_text_export_value(value: Any) -> str | None:
    if pd.isna(value):
        return None
    if isinstance(value, float) and value == int(value):
        # 避免 00254734 被读成 float 后丢失前导零（尽量保持原字符串）
        return str(int(value))
    s = str(value).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        return s[:-2]
    return s


def _excel_cell_value(value: Any) -> Any:
    if pd.isna(value):
        return None
    return value


def _display_width(text: str) -> float:
    """估算列宽：中文等宽字符按 2，英文按 1。"""
    width = 0.0
    for ch in text:
        width += 2.0 if ord(ch) > 127 else 1.0
    return width


def _write_merged_headers(ws, header_rows: list[list[Any]], ncol: int) -> int:
    """写入双行表头并还原合并单元格，返回数据起始行号。"""
    if len(header_rows) == 1:
        for ci in range(ncol):
            val = header_rows[0][ci] if ci < len(header_rows[0]) else None
            ws.cell(row=1, column=ci + 1, value=_excel_cell_value(val))
        return 2

    row0 = list(header_rows[0]) + [None] * max(0, ncol - len(header_rows[0]))
    row1 = list(header_rows[1]) + [None] * max(0, ncol - len(header_rows[1]))
    row0 = row0[:ncol]
    row1 = row1[:ncol]
    row0_ffill = pd.Series(row0).ffill().tolist()

    # 无子表头：第 1、2 行纵向合并
    for ci in range(ncol):
        if _cell_str(row1[ci]):
            continue
        val = row0[ci] if _cell_str(row0[ci]) else row0_ffill[ci]
        ws.cell(row=1, column=ci + 1, value=_excel_cell_value(val))
        ws.merge_cells(
            start_row=1,
            start_column=ci + 1,
            end_row=2,
            end_column=ci + 1,
        )

    # 有子表头：第 2 行写子表头，第 1 行按父标题横向合并
    ci = 0
    while ci < ncol:
        if not _cell_str(row1[ci]):
            ci += 1
            continue
        parent = _cell_str(row0_ffill[ci])
        start_ci = ci
        while ci < ncol and _cell_str(row1[ci]) and _cell_str(row0_ffill[ci]) == parent:
            ws.cell(row=2, column=ci + 1, value=_excel_cell_value(row1[ci]))
            ci += 1
        end_ci = ci - 1
        if not parent:
            continue
        ws.cell(row=1, column=start_ci + 1, value=parent)
        if end_ci > start_ci:
            ws.merge_cells(
                start_row=1,
                start_column=start_ci + 1,
                end_row=1,
                end_column=end_ci + 1,
            )

    return 3


def apply_worksheet_style(ws) -> None:
    """行高 30、Arial 10 号、居中、按内容自适应列宽。"""
    font = Font(name="Arial", size=FONT_SIZE)
    alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    col_max_width: dict[int, float] = {}

    for row in ws.iter_rows():
        row_idx = row[0].row
        ws.row_dimensions[row_idx].height = ROW_HEIGHT
        for cell in row:
            cell.font = font
            cell.alignment = alignment
            if cell.value is not None:
                col_idx = cell.column
                w = _display_width(str(cell.value))
                col_max_width[col_idx] = max(col_max_width.get(col_idx, 0.0), w)

    for col_idx, max_w in col_max_width.items():
        letter = get_column_letter(col_idx)
        width = min(max(max_w + COL_WIDTH_PADDING, MIN_COL_WIDTH), MAX_COL_WIDTH)
        ws.column_dimensions[letter].width = width


def _write_data_cell(ws, row: int, col: int, value: Any, kind: str) -> None:
    if pd.isna(value):
        ws.cell(row=row, column=col, value=None)
        return

    cell = ws.cell(row=row, column=col, value=_excel_cell_value(value))
    if kind == "text":
        cell.number_format = "@"
        cell.value = _to_text_export_value(value)
    elif kind == "number" and not isinstance(value, str):
        if isinstance(value, float) and value == int(value):
            cell.value = int(value)


def _populate_worksheet(
    ws,
    export_df: pd.DataFrame,
    col_kinds: dict[str, str],
    header_rows: list[list[Any]] | None = None,
) -> None:
    ncol = len(export_df.columns)
    columns = list(export_df.columns)

    if header_rows:
        start_row = _write_merged_headers(ws, header_rows, ncol)
    else:
        for ci, col_name in enumerate(columns, start=1):
            ws.cell(row=1, column=ci, value=col_name)
        start_row = 2

    for ri, row in enumerate(export_df.itertuples(index=False), start=start_row):
        for ci, val in enumerate(row, start=1):
            _write_data_cell(ws, ri, ci, val, col_kinds[columns[ci - 1]])

    apply_worksheet_style(ws)


def export_dataframe_to_excel(
    df: pd.DataFrame,
    path: str | Path,
    header_rows: list[list[Any]] | None = None,
    col_kinds: dict[str, str] | None = None,
) -> Path:
    export_df, col_kinds = prepare_dataframe_for_export(df, col_kinds)
    path = Path(path)

    wb = Workbook()
    ws = wb.active
    _populate_worksheet(ws, export_df, col_kinds, header_rows)
    return _save_workbook_safe(wb, path)


def sanitize_sheet_name(name: str, used: set[str]) -> str:
    s = INVALID_SHEET_CHARS.sub("_", str(name).strip())
    s = s.strip(" .") or "Sheet"
    if len(s) > MAX_SHEET_NAME_LEN:
        s = s[:MAX_SHEET_NAME_LEN]
    base = s
    n = 1
    while s in used:
        suffix = f"_{n}"
        s = (base[: MAX_SHEET_NAME_LEN - len(suffix)] + suffix) if len(base) + len(suffix) > MAX_SHEET_NAME_LEN else base + suffix
        n += 1
    used.add(s)
    return s


def make_timestamped_output_file(
    source_file: str | Path,
    base_dir: str | Path | None = None,
) -> Path:
    """生成单个输出工作簿路径。"""
    source_file = Path(source_file)
    parent = Path(base_dir) if base_dir else source_file.parent
    stem = sanitize_filename(source_file.stem, max_len=50)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return parent / f"{stem}_{OUTPUT_FOLDER_LABEL}_{ts}.xlsx"


def _save_workbook_safe(wb: Workbook, path: Path) -> Path:
    """先写入临时文件再替换，避免写入不完整。"""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(suffix=".xlsx", dir=path.parent)
    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        wb.save(tmp_path)
        if hasattr(wb, "close"):
            wb.close()
        os.replace(tmp_path, path)
        return path
    finally:
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass


def sanitize_filename(name: str, max_len: int = 80) -> str:
    s = INVALID_FILENAME_CHARS.sub("_", str(name).strip())
    s = s.strip(" .") or "未命名"
    if len(s) > max_len:
        s = s[:max_len]
    return s


def make_timestamped_output_dir(
    source_file: str | Path,
    base_dir: str | Path | None = None,
) -> Path:
    """生成「源文件名_拆分表格_时间戳」输出目录（每次拆分独立，无需清空旧文件）。"""
    source_file = Path(source_file)
    parent = Path(base_dir) if base_dir else source_file.parent
    stem = sanitize_filename(source_file.stem, max_len=60)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = parent / f"{stem}_{OUTPUT_FOLDER_LABEL}_{ts}"
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir


def _group_keys_and_frames(df: pd.DataFrame, split_columns: list[str]):
    if not split_columns:
        yield ("全部数据",), df
        return
    missing = [c for c in split_columns if c not in df.columns]
    if missing:
        raise KeyError(f"列不存在: {', '.join(missing)}")
    grouped = df.groupby(split_columns, dropna=False)
    for key, group_df in grouped:
        if not isinstance(key, tuple):
            key = (key,)
        yield key, group_df


def split_and_export(
    df: pd.DataFrame,
    split_columns: list[str],
    output_path: str | Path,
    header_rows: list[list[Any]] | None = None,
    col_kinds: dict[str, str] | None = None,
    mode: str = OUTPUT_MODE_FILES,
    on_progress: Callable[[int, int], None] | None = None,
) -> SplitExportResult:
    if mode == OUTPUT_MODE_WORKBOOK:
        return _split_to_workbook(df, split_columns, output_path, header_rows, col_kinds, on_progress)

    output_dir = Path(output_path)
    output_dir.mkdir(parents=True, exist_ok=True)

    groups = list(_group_keys_and_frames(df, split_columns))
    if not groups:
        raise ValueError("没有可写入的数据")
    total = len(groups)
    count = 0
    for key, group_df in groups:
        parts = [sanitize_filename(k if pd.notna(k) else "空值") for k in key]
        base_name = "_".join(parts)
        export_dataframe_to_excel(group_df, output_dir / f"{base_name}.xlsx", header_rows, col_kinds)
        count += 1
        if on_progress:
            on_progress(count, total)

    return SplitExportResult(count=count, output_path=output_dir, mode=OUTPUT_MODE_FILES)


def _split_to_workbook(
    df: pd.DataFrame,
    split_columns: list[str],
    output_file: str | Path,
    header_rows: list[list[Any]] | None = None,
    col_kinds: dict[str, str] | None = None,
    on_progress: Callable[[int, int], None] | None = None,
) -> SplitExportResult:
    output_file = Path(output_file)
    output_file.parent.mkdir(parents=True, exist_ok=True)

    groups = list(_group_keys_and_frames(df, split_columns))
    if not groups:
        raise ValueError("没有可写入的数据")
    total = len(groups)

    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    used_sheet_names: set[str] = set()
    count = 0

    for key, group_df in groups:
        parts = [sanitize_filename(k if pd.notna(k) else "空值", max_len=20) for k in key]
        sheet_title = sanitize_sheet_name("_".join(parts), used_sheet_names)
        export_df, kinds = prepare_dataframe_for_export(group_df, col_kinds)
        ws = wb.create_sheet(title=sheet_title)
        _populate_worksheet(ws, export_df, kinds, header_rows)
        count += 1
        if on_progress:
            on_progress(count, total)

    _save_workbook_safe(wb, output_file)
    return SplitExportResult(count=count, output_path=output_file, mode=OUTPUT_MODE_WORKBOOK)
