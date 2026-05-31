"""Excel 按列拆分与筛选 — CustomTkinter GUI（Element UI 风格）。"""
from __future__ import annotations

import os
import subprocess
import sys
import tkinter as tk
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog
from typing import Any

import customtkinter as ctk
import pandas as pd
from tkcalendar import Calendar

from excel_processor import (
    COLUMN_KIND_FROM_LABEL,
    COLUMN_KIND_LABELS,
    OUTPUT_FOLDER_LABEL,
    OUTPUT_MODE_FILES,
    OUTPUT_MODE_WORKBOOK,
    SplitExportResult,
    apply_filters,
    detect_all_column_kinds,
    infer_column_kind,
    list_sheet_names,
    make_timestamped_output_dir,
    make_timestamped_output_file,
    read_excel_with_headers,
    split_and_export,
)
from ui_theme import (
    BG_PAGE,
    BORDER,
    BORDER_LIGHT,
    CARD_BG,
    DANGER,
    FONT_BODY,
    FONT_HINT,
    FONT_SECTION,
    FONT_SMALL,
    FONT_TITLE,
    TEXT_PLACEHOLDER,
    HEADER_BG,
    HEADER_FG,
    PRIMARY,
    PRIMARY_HOVER,
    SUCCESS,
    TEXT_PRIMARY,
    TEXT_REGULAR,
    TEXT_SECONDARY,
    apply_window_icon,
    card,
    configure_app,
    danger_button,
    hint_label,
    primary_button,
    scroll_frame_to_top,
    secondary_button,
    section_title,
    shorten_path,
)

KIND_OPTIONS = list(COLUMN_KIND_LABELS.values())
TYPE_TABS_UI = (
    ("all", "全部"),
    ("text", "文本"),
    ("date", "日期"),
    ("number", "数字"),
)


def _parse_date_str(value: str) -> date | None:
    text = (value or "").strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y/%m/%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    try:
        return datetime.strptime(text.split()[0], "%Y-%m-%d").date()
    except ValueError:
        return None


def _maximize_window(root: ctk.CTk) -> None:
    try:
        root.state("zoomed")
        return
    except tk.TclError:
        pass
    try:
        root.attributes("-zoomed", True)
        return
    except tk.TclError:
        pass
    w = root.winfo_screenwidth()
    h = root.winfo_screenheight()
    root.geometry(f"{w}x{h}+0+0")


class MessageDialog(ctk.CTkToplevel):
    """Element 风格提示框。"""

    def __init__(
        self,
        parent: ctk.CTk,
        title: str,
        message: str,
        level: str = "info",
    ):
        super().__init__(parent)
        self.title(title)
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)
        apply_window_icon(self)

        accent = {"info": PRIMARY, "warning": "#E6A23C", "error": DANGER}.get(level, PRIMARY)
        top_bar = ctk.CTkFrame(self, fg_color=accent, height=4, corner_radius=0)
        top_bar.pack(fill="x")

        body = ctk.CTkFrame(self, fg_color=CARD_BG, corner_radius=0)
        body.pack(fill="both", expand=True, padx=0, pady=0)

        ctk.CTkLabel(
            body,
            text=message,
            font=FONT_BODY,
            text_color=TEXT_PRIMARY,
            justify="left",
            wraplength=400,
        ).pack(anchor="w", padx=24, pady=(20, 16))

        primary_button(body, text="确定", command=self.destroy, width=88).pack(
            anchor="e", padx=24, pady=(0, 20)
        )

        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")


def _show_info(parent: ctk.CTk, title: str, message: str) -> None:
    MessageDialog(parent, title, message, "info")


def _show_warning(parent: ctk.CTk, title: str, message: str) -> None:
    MessageDialog(parent, title, message, "warning")


def _show_error(parent: ctk.CTk, title: str, message: str) -> None:
    MessageDialog(parent, title, message, "error")


_CALENDAR_STYLE = {
    "background": CARD_BG,
    "foreground": TEXT_PRIMARY,
    "headersbackground": PRIMARY,
    "headersforeground": "#FFFFFF",
    "selectbackground": PRIMARY,
    "selectforeground": "#FFFFFF",
    "normalbackground": CARD_BG,
    "normalforeground": TEXT_PRIMARY,
    "weekendbackground": "#F5F7FA",
    "weekendforeground": TEXT_REGULAR,
    "othermonthforeground": TEXT_PLACEHOLDER,
    "othermonthbackground": BORDER_LIGHT,
    "bordercolor": BORDER,
    "borderwidth": 0,
    "font": FONT_SMALL,
    "headersfont": (FONT_SMALL[0], FONT_SMALL[1], "bold"),
}


def _center_toplevel(win: ctk.CTkToplevel, parent: ctk.CTk) -> None:
    win.update_idletasks()
    x = parent.winfo_x() + (parent.winfo_width() - win.winfo_width()) // 2
    y = parent.winfo_y() + (parent.winfo_height() - win.winfo_height()) // 2
    win.geometry(f"+{max(x, 0)}+{max(y, 0)}")


class DatePickerDialog(ctk.CTkToplevel):
    """Element 风格日期选择弹窗。"""

    def __init__(
        self,
        parent: ctk.CTk,
        title: str,
        initial: date,
        on_confirm,
        on_clear=None,
    ):
        super().__init__(parent)
        self.title(title)
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)
        apply_window_icon(self)
        self.configure(fg_color=CARD_BG)

        ctk.CTkFrame(self, fg_color=PRIMARY, height=4, corner_radius=0).pack(fill="x")

        header = ctk.CTkFrame(self, fg_color=PRIMARY, corner_radius=0, height=44)
        header.pack(fill="x")
        header.pack_propagate(False)
        ctk.CTkLabel(
            header,
            text=title,
            font=FONT_SECTION,
            text_color="#FFFFFF",
        ).pack(side="left", padx=20, pady=8)

        body = ctk.CTkFrame(self, fg_color=CARD_BG, corner_radius=0)
        body.pack(fill="both", expand=True)

        cal_wrap = ctk.CTkFrame(
            body,
            fg_color=CARD_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER,
        )
        cal_wrap.pack(padx=20, pady=(16, 12))

        cal_host = tk.Frame(cal_wrap, bg=CARD_BG, bd=0, highlightthickness=0)
        cal_host.pack(padx=8, pady=8)

        self._cal = Calendar(
            cal_host,
            selectmode="day",
            year=initial.year,
            month=initial.month,
            day=initial.day,
            date_pattern="yyyy-mm-dd",
            showweeknumbers=False,
            **_CALENDAR_STYLE,
        )
        self._cal.pack()

        btn_row = ctk.CTkFrame(body, fg_color="transparent")
        btn_row.pack(pady=(0, 18))

        def confirm() -> None:
            on_confirm(self._cal.selection_get())
            self.destroy()

        def clear() -> None:
            if on_clear:
                on_clear()
            self.destroy()

        primary_button(btn_row, text="确定", command=confirm, width=80).pack(side="left", padx=6)
        secondary_button(btn_row, text="清空", command=clear, width=80).pack(side="left", padx=6)
        secondary_button(btn_row, text="取消", command=self.destroy, width=80).pack(side="left", padx=6)

        _center_toplevel(self, parent)


class FilterEditor:
    """单行内嵌筛选条件（列名右侧）。"""

    def __init__(
        self,
        parent_row: ctk.CTkFrame,
        df: pd.DataFrame,
        column: str,
        get_col_kind,
        on_remove=None,
    ):
        self.col = column
        self.df = df
        self.get_col_kind = get_col_kind
        self.row = parent_row
        self.cond_frame = ctk.CTkFrame(parent_row, fg_color="transparent")
        self.cond_frame.pack(side="left", fill="x", expand=True, padx=(4, 4), pady=6)
        if on_remove:
            secondary_button(
                parent_row,
                text="移除",
                width=52,
                height=28,
                command=on_remove,
            ).pack(side="right", padx=(0, 10), pady=6)
        self._rebuild_cond_widgets()

    def _open_date_picker(self, entry: ctk.CTkEntry, title: str = "选择日期") -> None:
        parent = self.row.winfo_toplevel()
        initial = _parse_date_str(entry.get()) or date.today()

        def on_confirm(picked: date) -> None:
            entry.delete(0, "end")
            entry.insert(0, picked.strftime("%Y-%m-%d"))

        def on_clear() -> None:
            entry.delete(0, "end")

        DatePickerDialog(parent, title, initial, on_confirm, on_clear)

    def _build_date_widgets(self) -> None:
        ctk.CTkLabel(self.cond_frame, text="起始", font=FONT_SMALL, text_color=TEXT_REGULAR).pack(
            side="left"
        )
        self.start_entry = ctk.CTkEntry(self.cond_frame, width=88, font=FONT_SMALL)
        self.start_entry.pack(side="left", padx=4)
        secondary_button(
            self.cond_frame,
            text="选择",
            width=56,
            height=28,
            command=lambda: self._open_date_picker(self.start_entry, "选择起始日期"),
        ).pack(side="left", padx=(0, 8))

        ctk.CTkLabel(self.cond_frame, text="截止", font=FONT_SMALL, text_color=TEXT_REGULAR).pack(
            side="left"
        )
        self.end_entry = ctk.CTkEntry(self.cond_frame, width=88, font=FONT_SMALL)
        self.end_entry.pack(side="left", padx=4)
        secondary_button(
            self.cond_frame,
            text="选择",
            width=56,
            height=28,
            command=lambda: self._open_date_picker(self.end_entry, "选择截止日期"),
        ).pack(side="left")

    def _rebuild_cond_widgets(self):
        for w in self.cond_frame.winfo_children():
            w.destroy()
        for attr in ("start_entry", "end_entry", "min_entry", "max_entry", "text_entry", "kind"):
            if hasattr(self, attr):
                delattr(self, attr)

        col = self.col
        if col not in self.df.columns:
            return
        kind = self.get_col_kind(col)
        self.kind = kind

        if kind == "date":
            self._build_date_widgets()
        elif kind == "number":
            ctk.CTkLabel(self.cond_frame, text="最小", font=FONT_SMALL, text_color=TEXT_REGULAR).pack(
                side="left"
            )
            self.min_entry = ctk.CTkEntry(self.cond_frame, width=72, font=FONT_SMALL)
            self.min_entry.pack(side="left", padx=4)
            ctk.CTkLabel(self.cond_frame, text="最大", font=FONT_SMALL, text_color=TEXT_REGULAR).pack(
                side="left"
            )
            self.max_entry = ctk.CTkEntry(self.cond_frame, width=72, font=FONT_SMALL)
            self.max_entry.pack(side="left", padx=4)
        else:
            ctk.CTkLabel(
                self.cond_frame, text="包含", font=FONT_SMALL, text_color=TEXT_REGULAR
            ).pack(side="left")
            self.text_entry = ctk.CTkEntry(self.cond_frame, font=FONT_SMALL, placeholder_text="输入关键词…")
            self.text_entry.pack(side="left", fill="x", expand=True, padx=(4, 4))

    def to_filter_dict(self) -> dict[str, Any] | None:
        col = self.col
        if not col or col not in self.df.columns:
            return None
        if hasattr(self, "start_entry"):
            start = self.start_entry.get().strip()
            end = self.end_entry.get().strip()
            if not start and not end:
                return None
            return {"column": col, "kind": "date", "start": start or None, "end": end or None}
        if hasattr(self, "min_entry"):
            lo = self.min_entry.get().strip()
            hi = self.max_entry.get().strip()
            if not lo and not hi:
                return None
            return {"column": col, "kind": "number", "min": lo or None, "max": hi or None}
        if hasattr(self, "text_entry"):
            text = self.text_entry.get().strip()
            if not text:
                return None
            return {"column": col, "kind": "text", "text": text}
        return None


class FilterPanel:
    """筛选列列表，条件内嵌在列名后。"""

    def __init__(
        self,
        parent: ctk.CTkFrame,
        height: int,
        get_col_kind,
        on_remove=None,
    ):
        self._items: list[str] = []
        self._selected: str | None = None
        self._row_widgets: dict[str, ctk.CTkFrame] = {}
        self.editors: dict[str, FilterEditor] = {}
        self.get_col_kind = get_col_kind
        self.on_remove = on_remove
        self.df: pd.DataFrame | None = None
        self.scroll = ctk.CTkScrollableFrame(
            parent,
            fg_color=BORDER_LIGHT,
            height=height,
            corner_radius=6,
            border_width=1,
            border_color=BORDER,
        )
        self.scroll.grid(row=0, column=0, sticky="nsew")
        parent.grid_rowconfigure(0, weight=1)
        parent.grid_columnconfigure(0, weight=1)

    def clear(self) -> None:
        for w in self.scroll.winfo_children():
            w.destroy()
        self._items.clear()
        self._selected = None
        self._row_widgets.clear()
        self.editors.clear()

    def get_all(self) -> list[str]:
        return list(self._items)

    def add(self, df: pd.DataFrame, col: str) -> None:
        self.df = df
        if col in self._items:
            self._select(col)
            return
        self._items.append(col)
        row = ctk.CTkFrame(self.scroll, fg_color=CARD_BG, corner_radius=4)
        row.pack(fill="x", pady=2, padx=4)
        self._row_widgets[col] = row

        ctk.CTkLabel(
            row,
            text=col,
            font=FONT_SMALL,
            text_color=TEXT_PRIMARY,
            anchor="w",
            width=72,
        ).pack(side="left", padx=(10, 0), pady=6)

        remove_cb = (lambda c=col: self.on_remove(c)) if self.on_remove else None
        self.editors[col] = FilterEditor(row, df, col, self.get_col_kind, on_remove=remove_cb)
        self._select(col)

    def _select(self, col: str | None) -> None:
        self._selected = col
        for name, row in self._row_widgets.items():
            if col and name == col:
                row.configure(fg_color="#ECF5FF", border_width=1, border_color=PRIMARY)
            else:
                row.configure(fg_color=CARD_BG, border_width=0)

    def remove_col(self, col: str) -> None:
        if col in self._row_widgets:
            self._row_widgets[col].destroy()
            del self._row_widgets[col]
        if col in self.editors:
            del self.editors[col]
        if col in self._items:
            self._items.remove(col)
        if self._selected == col:
            self._selected = self._items[-1] if self._items else None
            self._select(self._selected)

    def remove_selected(self) -> str | None:
        if not self._selected:
            return None
        col = self._selected
        self.remove_col(col)
        return col


class SplitColumnPanel:
    """已选拆分列列表。"""

    def __init__(self, parent: ctk.CTkFrame, height: int = 200, on_select=None):
        self._items: list[str] = []
        self._selected: str | None = None
        self._row_widgets: dict[str, ctk.CTkFrame] = {}
        self.on_select = on_select
        self.scroll = ctk.CTkScrollableFrame(
            parent,
            fg_color=BORDER_LIGHT,
            height=height,
            corner_radius=6,
            border_width=1,
            border_color=BORDER,
        )
        self.scroll.grid(row=0, column=0, sticky="nsew")
        parent.grid_rowconfigure(0, weight=1)
        parent.grid_columnconfigure(0, weight=1)

    def clear(self) -> None:
        for w in self.scroll.winfo_children():
            w.destroy()
        self._items.clear()
        self._selected = None
        self._row_widgets.clear()

    def get_all(self) -> list[str]:
        return list(self._items)

    def add(self, col: str) -> None:
        if col in self._items:
            self._select(col)
            return
        self._items.append(col)
        row = ctk.CTkFrame(self.scroll, fg_color=CARD_BG, corner_radius=4, height=32)
        row.pack(fill="x", pady=2, padx=4)
        row.pack_propagate(False)
        self._row_widgets[col] = row

        def on_click(_e=None, c=col):
            self._select(c)

        lbl = ctk.CTkLabel(
            row,
            text=col,
            font=FONT_SMALL,
            text_color=TEXT_PRIMARY,
            anchor="w",
        )
        lbl.pack(side="left", fill="x", expand=True, padx=10)
        lbl.bind("<Button-1>", on_click)
        row.bind("<Button-1>", on_click)
        self._select(col)

    def _select(self, col: str | None) -> None:
        self._selected = col
        for name, row in self._row_widgets.items():
            if col and name == col:
                row.configure(fg_color="#ECF5FF", border_width=1, border_color=PRIMARY)
            else:
                row.configure(fg_color=CARD_BG, border_width=0)
        if self.on_select:
            self.on_select(col)

    def remove_col(self, col: str) -> None:
        if col in self._row_widgets:
            self._row_widgets[col].destroy()
            del self._row_widgets[col]
        if col in self._items:
            self._items.remove(col)
        if self._selected == col:
            self._selected = self._items[-1] if self._items else None
            self._select(self._selected)

    def remove_selected(self) -> str | None:
        if not self._selected:
            return None
        col = self._selected
        self.remove_col(col)
        if self._selected:
            self._select(self._selected)
        return col


class ExcelSplitApp:
    def __init__(self):
        configure_app()
        self.root = ctk.CTk()
        self.root.title("Excel 表格拆分工具")
        self.root.minsize(860, 720)
        self.root.configure(fg_color=BG_PAGE)
        apply_window_icon(self.root)

        self.file_path: str | None = None
        self.df: pd.DataFrame | None = None
        self.header_rows: list | None = None
        self.filter_panel: FilterPanel | None = None
        self.auto_col_kinds: dict[str, str] = {}
        self.col_kind_vars: dict[str, tk.StringVar] = {}
        self.col_type_rows: dict[str, ctk.CTkFrame] = {}
        self.type_tab_buttons: dict[str, ctk.CTkButton] = {}
        self.kind_combos: dict[str, ctk.CTkComboBox] = {}

        self._build_ui()
        self.root.after(100, lambda: _maximize_window(self.root))

    def _build_ui(self):
        # 顶栏
        header = ctk.CTkFrame(self.root, fg_color=HEADER_BG, corner_radius=0, height=56)
        header.pack(fill="x")
        header.pack_propagate(False)
        ctk.CTkLabel(
            header,
            text="Excel 表格拆分工具",
            font=FONT_TITLE,
            text_color=HEADER_FG,
        ).pack(side="left", padx=24, pady=12)
        ctk.CTkLabel(
            header,
            text="按列拆分 · 筛选 · 多文件或单工作簿导出",
            font=FONT_HINT,
            text_color="#E8F4FF",
        ).pack(side="left", padx=8)

        # 底部操作区（先 pack 到底部，避免挤占中间区域）
        bottom_card = card(self.root)
        bottom_card.pack(side="bottom", fill="x", padx=16, pady=(0, 16))

        # 中间主内容可滚动，确保「筛选条件」等区块不会被挤出视口
        self.main_scroll = ctk.CTkScrollableFrame(
            self.root,
            fg_color=BG_PAGE,
            corner_radius=0,
            scrollbar_button_color=BORDER,
            scrollbar_button_hover_color=TEXT_SECONDARY,
        )
        self.main_scroll.pack(fill="both", expand=True, padx=16, pady=12)
        main = self.main_scroll

        mode_inner = ctk.CTkFrame(bottom_card, fg_color="transparent")
        mode_inner.pack(fill="x", padx=16, pady=(14, 6))
        ctk.CTkLabel(
            mode_inner, text="输出方式", font=FONT_SMALL, text_color=TEXT_REGULAR, width=72
        ).pack(side="left")
        self.output_mode_var = tk.StringVar(value=OUTPUT_MODE_FILES)
        ctk.CTkRadioButton(
            mode_inner,
            text="多个文件（文件夹）",
            variable=self.output_mode_var,
            value=OUTPUT_MODE_FILES,
            font=FONT_SMALL,
            command=self._refresh_output_hint,
        ).pack(side="left", padx=(8, 16))
        ctk.CTkRadioButton(
            mode_inner,
            text="单个工作簿（多工作表）",
            variable=self.output_mode_var,
            value=OUTPUT_MODE_WORKBOOK,
            font=FONT_SMALL,
            command=self._refresh_output_hint,
        ).pack(side="left")

        out_inner = ctk.CTkFrame(bottom_card, fg_color="transparent")
        out_inner.pack(fill="x", padx=16, pady=(0, 10))
        ctk.CTkLabel(
            out_inner, text="输出位置", font=FONT_SMALL, text_color=TEXT_REGULAR, width=72
        ).pack(side="left", anchor="n", pady=6)
        self.out_entry = ctk.CTkEntry(
            out_inner, font=FONT_SMALL, state="readonly", text_color=TEXT_SECONDARY
        )
        self.out_entry.pack(side="left", fill="x", expand=True, padx=(4, 0))

        action = ctk.CTkFrame(bottom_card, fg_color="transparent")
        action.pack(fill="x", padx=16, pady=(0, 6))
        action_row = ctk.CTkFrame(action, fg_color="transparent")
        action_row.pack(fill="x")
        self.run_btn = primary_button(
            action_row, text="开始拆分", command=self._run_split, width=120, height=38
        )
        self.run_btn.pack(side="left")
        self.status_label = ctk.CTkLabel(
            action_row,
            text="请选择 Excel 文件",
            font=FONT_SMALL,
            text_color=TEXT_SECONDARY,
        )
        self.status_label.pack(side="left", padx=20)
        self.progress_bar = ctk.CTkProgressBar(action, height=8, progress_color=PRIMARY)
        self.progress_bar.pack(fill="x", pady=(8, 0))
        self.progress_bar.set(0)
        self.progress_bar.pack_forget()

        # 文件 + 工作表（同一行）
        source_card = card(main)
        source_card.pack(fill="x", pady=(0, 10))
        section_title(source_card, "Excel 文件").pack(anchor="w", padx=16, pady=(14, 8))

        source_row = ctk.CTkFrame(source_card, fg_color="transparent")
        source_row.pack(fill="x", padx=16, pady=(0, 4))

        file_part = ctk.CTkFrame(source_row, fg_color="transparent")
        file_part.pack(side="left", padx=(0, 24))
        ctk.CTkLabel(file_part, text="文件", font=FONT_SMALL, text_color=TEXT_REGULAR, width=36).pack(
            side="left"
        )
        self.path_entry = ctk.CTkEntry(
            file_part,
            width=200,
            font=FONT_SMALL,
            placeholder_text="未选择文件",
            state="readonly",
        )
        self.path_entry.pack(side="left", padx=(4, 8))
        primary_button(file_part, text="选择文件", command=self._browse_file, width=92).pack(side="left")

        sheet_part = ctk.CTkFrame(source_row, fg_color="transparent")
        sheet_part.pack(side="left")
        ctk.CTkLabel(sheet_part, text="工作表", font=FONT_SMALL, text_color=TEXT_REGULAR, width=48).pack(
            side="left"
        )
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ctk.CTkComboBox(
            sheet_part,
            variable=self.sheet_var,
            values=[],
            width=180,
            font=FONT_SMALL,
            command=lambda _: self._reload_sheet(),
        )
        self.sheet_combo.pack(side="left", padx=(4, 8))
        secondary_button(sheet_part, text="重新加载", command=self._reload_sheet, width=88).pack(side="left")

        self.path_hint_label = ctk.CTkLabel(
            source_card,
            text="",
            font=FONT_HINT,
            text_color=TEXT_SECONDARY,
            anchor="w",
        )
        self.path_hint_label.pack(fill="x", padx=16, pady=(0, 12))

        # 按列拆分 | 筛选条件 — 上排左右对称
        _PAD = 16
        _HINT_H = 40
        _TOOLBAR_H = 34
        _CTRL_H = 34
        _LIST_H = 200

        def _sym_panel(parent: ctk.CTkFrame, column: int) -> ctk.CTkFrame:
            p = card(parent)
            p.grid(row=0, column=column, sticky="nsew", padx=(0, 6) if column == 0 else (6, 0))
            for r in range(4):
                p.grid_rowconfigure(r, weight=0)
            p.grid_rowconfigure(4, weight=1)
            p.grid_columnconfigure(0, weight=1)
            return p

        def _sym_hint_row(panel: ctk.CTkFrame, row: int, text: str) -> None:
            wrap = ctk.CTkFrame(panel, fg_color="transparent", height=_HINT_H)
            wrap.grid(row=row, column=0, sticky="ew", padx=_PAD, pady=(0, 6))
            wrap.grid_propagate(False)
            hint_label(wrap, text).pack(anchor="nw", fill="x")

        def _sym_toolbar_row(panel: ctk.CTkFrame, row: int) -> ctk.CTkFrame:
            wrap = ctk.CTkFrame(panel, fg_color="transparent", height=_TOOLBAR_H)
            wrap.grid(row=row, column=0, sticky="ew", padx=_PAD, pady=(0, 6))
            wrap.grid_propagate(False)
            return wrap

        def _sym_ctrl_row(panel: ctk.CTkFrame, row: int) -> ctk.CTkFrame:
            wrap = ctk.CTkFrame(panel, fg_color="transparent", height=_CTRL_H)
            wrap.grid(row=row, column=0, sticky="ew", padx=_PAD, pady=(0, 6))
            wrap.grid_propagate(False)
            return wrap

        upper_row = ctk.CTkFrame(main, fg_color="transparent")
        upper_row.pack(fill="both", expand=True, pady=(0, 10))
        upper_row.grid_columnconfigure(0, weight=1, uniform="upper")
        upper_row.grid_columnconfigure(1, weight=1, uniform="upper")
        upper_row.grid_rowconfigure(0, weight=1)

        split_card = _sym_panel(upper_row, 0)
        section_title(split_card, "按列拆分").grid(row=0, column=0, sticky="w", padx=_PAD, pady=(14, 4))
        _sym_hint_row(split_card, 1, "可多列组合；文件夹→多文件，工作簿→多工作表")

        split_toolbar = _sym_toolbar_row(split_card, 2)
        hint_label(split_toolbar, "从下拉选择列后点击「添加」").pack(side="left", pady=2)

        split_ctrl = _sym_ctrl_row(split_card, 3)
        ctk.CTkLabel(split_ctrl, text="拆分列", font=FONT_SMALL, text_color=TEXT_REGULAR, width=48).pack(
            side="left", pady=2
        )
        self.split_col_var = tk.StringVar()
        self.split_col_combo = ctk.CTkComboBox(
            split_ctrl, variable=self.split_col_var, values=[], width=120, font=FONT_SMALL
        )
        self.split_col_combo.pack(side="left", fill="x", expand=True, padx=(4, 8), pady=2)
        secondary_button(split_ctrl, text="添加", command=self._add_split_column, width=64).pack(
            side="left", padx=2, pady=2
        )
        secondary_button(split_ctrl, text="移除", command=self._remove_split_column, width=64).pack(
            side="left", padx=2, pady=2
        )

        split_list_wrap = ctk.CTkFrame(split_card, fg_color="transparent")
        split_list_wrap.grid(row=4, column=0, sticky="nsew", padx=_PAD, pady=(0, 14))
        split_list_wrap.grid_rowconfigure(0, weight=1)
        split_list_wrap.grid_columnconfigure(0, weight=1)
        self.split_panel = SplitColumnPanel(split_list_wrap)

        self.filter_card = _sym_panel(upper_row, 1)
        section_title(self.filter_card, "筛选条件").grid(
            row=0, column=0, sticky="w", padx=_PAD, pady=(14, 4)
        )
        _sym_hint_row(self.filter_card, 1, "添加后于列名右侧设置条件（依列类型显示）")

        filter_toolbar = _sym_toolbar_row(self.filter_card, 2)
        hint_label(filter_toolbar, "从下拉选择列后点击「添加」").pack(side="left", pady=2)

        filter_ctrl = _sym_ctrl_row(self.filter_card, 3)
        ctk.CTkLabel(filter_ctrl, text="筛选列", font=FONT_SMALL, text_color=TEXT_REGULAR, width=48).pack(
            side="left", pady=2
        )
        self.filter_col_var = tk.StringVar()
        self.filter_col_combo = ctk.CTkComboBox(
            filter_ctrl, variable=self.filter_col_var, values=[], width=120, font=FONT_SMALL
        )
        self.filter_col_combo.pack(side="left", fill="x", expand=True, padx=(4, 8), pady=2)
        secondary_button(filter_ctrl, text="添加", command=self._add_filter, width=64).pack(
            side="left", padx=2, pady=2
        )
        secondary_button(filter_ctrl, text="移除", command=self._remove_filter, width=64).pack(
            side="left", padx=2, pady=2
        )

        filter_list_wrap = ctk.CTkFrame(self.filter_card, fg_color="transparent")
        filter_list_wrap.grid(row=4, column=0, sticky="nsew", padx=_PAD, pady=(0, 14))
        filter_list_wrap.grid_rowconfigure(0, weight=1)
        filter_list_wrap.grid_columnconfigure(0, weight=1)
        self.filter_panel = FilterPanel(
            filter_list_wrap,
            height=_LIST_H,
            get_col_kind=self._get_col_kind,
            on_remove=self._remove_filter_for_column,
        )

        # 列类型 — 下排全宽
        self.type_card = card(main)
        self.type_card.pack(fill="x", pady=(0, 10))
        type_card = self.type_card
        section_title(type_card, "列类型").pack(anchor="w", padx=_PAD, pady=(14, 4))
        hint_label(type_card, "加载后自动判断，可在「修正为」中修改；影响筛选与导出格式").pack(
            anchor="w", padx=_PAD, pady=(0, 8)
        )

        type_toolbar = ctk.CTkFrame(type_card, fg_color="transparent")
        type_toolbar.pack(fill="x", padx=_PAD, pady=(0, 6))
        secondary_button(
            type_toolbar, text="全部恢复自动判断", command=self._reset_column_kinds, width=140
        ).pack(side="left")
        hint_label(type_toolbar, "日期 → YYYY/MM/DD；数字 → 数值；文本 → 编号等").pack(
            side="left", padx=12
        )

        tab_bar = ctk.CTkFrame(type_card, fg_color="transparent")
        tab_bar.pack(fill="x", padx=_PAD, pady=(0, 6))
        self.type_tab_var = tk.StringVar(value="all")
        for key, label in TYPE_TABS_UI:
            btn = ctk.CTkButton(
                tab_bar,
                text=label,
                font=FONT_SMALL,
                width=88,
                height=28,
                fg_color=BORDER_LIGHT,
                hover_color=BORDER,
                text_color=TEXT_REGULAR,
                command=lambda k=key: self._on_type_tab(k),
            )
            btn.pack(side="left", padx=3)
            self.type_tab_buttons[key] = btn
        self._highlight_type_tab("all")

        self.type_scroll = ctk.CTkScrollableFrame(
            type_card,
            fg_color=BORDER_LIGHT,
            height=160,
            corner_radius=6,
            border_width=1,
            border_color=BORDER,
        )
        self.type_scroll.pack(fill="x", padx=_PAD, pady=(0, 14))

        header_row = ctk.CTkFrame(self.type_scroll, fg_color="transparent")
        header_row.pack(fill="x", pady=(4, 6))
        ctk.CTkLabel(header_row, text="列名", anchor="w", font=FONT_SMALL, text_color=TEXT_SECONDARY).pack(
            side="left", fill="x", expand=True, padx=8
        )
        ctk.CTkLabel(header_row, text="自动判断", width=64, font=FONT_SMALL, text_color=TEXT_SECONDARY).pack(
            side="left"
        )
        ctk.CTkLabel(header_row, text="修正为", width=72, font=FONT_SMALL, text_color=TEXT_SECONDARY).pack(
            side="left", padx=(0, 8)
        )
        self.type_list_container = ctk.CTkFrame(self.type_scroll, fg_color="transparent")
        self.type_list_container.pack(fill="x")

        self._refresh_output_hint()

    def _on_type_tab(self, key: str) -> None:
        self.type_tab_var.set(key)
        self._highlight_type_tab(key)
        self._refresh_type_tab_view()
        self.root.after(10, self._scroll_to_type_section)

    def _scroll_to_type_section(self) -> None:
        """切换列类型 Tab 时滚到「列类型」区块，并重置列表滚动位置。"""
        scroll_frame_to_top(self.type_scroll)
        if hasattr(self, "type_card"):
            self._scroll_main_to_widget(self.type_card)

    def _scroll_panels_to_top(self) -> None:
        scroll_frame_to_top(self.type_scroll)
        if self.filter_panel:
            scroll_frame_to_top(self.filter_panel.scroll)
        if hasattr(self.split_panel, "scroll"):
            scroll_frame_to_top(self.split_panel.scroll)

    def _scroll_main_to_widget(self, widget, padding: int = 4) -> None:
        """将主滚动区滚到指定区块顶部。"""
        try:
            canvas = self.main_scroll._parent_canvas
            inner = self.main_scroll._parent_frame
            canvas.update_idletasks()
            inner.update_idletasks()
            inner_h = inner.winfo_reqheight()
            viewport_h = max(canvas.winfo_height(), 1)
            if inner_h <= viewport_h:
                canvas.yview_moveto(0)
                return
            y = widget.winfo_rooty() - inner.winfo_rooty()
            target = max(0, y - padding)
            max_scroll = inner_h - viewport_h
            canvas.yview_moveto(max(0, min(1, target / max_scroll)))
        except Exception:
            pass

    def _highlight_type_tab(self, active: str) -> None:
        for key, btn in self.type_tab_buttons.items():
            if key == active:
                btn.configure(fg_color=PRIMARY, hover_color=PRIMARY_HOVER, text_color="#FFFFFF")
            else:
                btn.configure(fg_color=BORDER_LIGHT, hover_color=BORDER, text_color=TEXT_REGULAR)

    def _set_status(self, text: str) -> None:
        self.status_label.configure(text=text)

    def _show_progress(self, visible: bool) -> None:
        if visible:
            self.progress_bar.pack(fill="x", pady=(8, 0))
        else:
            self.progress_bar.pack_forget()
            self.progress_bar.set(0)

    def _report_split_progress(self, current: int, total: int) -> None:
        if total <= 0:
            return
        self.progress_bar.set(current / total)
        self._set_status(f"正在处理… ({current}/{total})")
        self.root.update_idletasks()

    def _set_path_display(self, path: str | None) -> None:
        self.path_entry.configure(state="normal")
        self.path_entry.delete(0, "end")
        if path:
            self.path_entry.insert(0, Path(path).name)
            self.path_hint_label.configure(text=shorten_path(path, 72))
        else:
            self.path_hint_label.configure(text="")
        self.path_entry.configure(state="readonly")

    def _set_out_hint(self, text: str) -> None:
        self.out_entry.configure(state="normal")
        self.out_entry.delete(0, "end")
        display = text
        if len(text) > 56 and ("\\" in text or "/" in text):
            display = shorten_path(text.replace("{", "").replace("}", ""), 56)
            if "（" in text:
                display = display + " " + text[text.find("（") :]
        self.out_entry.insert(0, display)
        self.out_entry.configure(state="readonly")

    def _reset_split_and_filters(self) -> None:
        self.split_panel.clear()
        self.split_col_var.set("")
        self._clear_filters()

    def _browse_file(self):
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel", "*.xls *.xlsx"), ("所有文件", "*.*")],
        )
        if path:
            self._reset_split_and_filters()
            self.file_path = path
            self._set_path_display(path)
            self._load_sheets()

    def _load_sheets(self):
        if not self.file_path:
            return
        try:
            names = list_sheet_names(self.file_path)
            self.sheet_combo.configure(values=names)
            if names:
                self.sheet_var.set(names[0])
            self._reload_sheet()
        except Exception as e:
            _show_error(self.root, "错误", f"无法读取文件:\n{e}")

    def _reload_sheet(self):
        if not self.file_path:
            return
        sheet = self.sheet_var.get() or 0
        try:
            self._set_status("正在加载…")
            self.root.update_idletasks()
            sheet_data = read_excel_with_headers(self.file_path, sheet_name=sheet)
            self.df = sheet_data.df
            self.header_rows = sheet_data.header_rows
            cols = list(self.df.columns)
            self.split_col_combo.configure(values=cols)
            self.filter_col_combo.configure(values=cols)
            self.split_col_var.set("")
            self.filter_col_var.set("")
            self._refresh_output_hint()
            self._reset_split_and_filters()
            self._rebuild_column_type_panel()
            self._set_status(f"已加载 {len(self.df)} 行，{len(cols)} 列 — 请核对列类型")
        except Exception as e:
            _show_error(self.root, "错误", f"加载失败:\n{e}")
            self._set_status("加载失败")

    def _columns(self) -> list[str]:
        return list(self.df.columns) if self.df is not None else []

    def _get_col_kind(self, col: str) -> str:
        if col in self.col_kind_vars:
            label = self.col_kind_vars[col].get()
            return COLUMN_KIND_FROM_LABEL.get(label, "text")
        if self.df is not None and col in self.df.columns:
            return infer_column_kind(self.df[col], col)
        return "text"

    def _get_col_kinds_dict(self) -> dict[str, str]:
        return {col: self._get_col_kind(col) for col in self._columns()}

    def _column_matches_type_tab(self, col: str, tab_key: str) -> bool:
        if tab_key == "all":
            return True
        return self._get_col_kind(col) == tab_key

    def _type_tab_counts(self) -> dict[str, int]:
        counts = {"all": 0, "text": 0, "date": 0, "number": 0}
        for col in self._columns():
            counts["all"] += 1
            counts[self._get_col_kind(col)] += 1
        return counts

    def _update_type_tab_labels(self) -> None:
        counts = self._type_tab_counts()
        active = self.type_tab_var.get()
        for key, label in TYPE_TABS_UI:
            btn = self.type_tab_buttons.get(key)
            if btn:
                btn.configure(text=f"{label} ({counts[key]})")
        self._highlight_type_tab(active)

    def _refresh_type_tab_view(self) -> None:
        tab = self.type_tab_var.get()
        for col, row in self.col_type_rows.items():
            if self._column_matches_type_tab(col, tab):
                row.pack(fill="x", pady=2)
            else:
                row.pack_forget()
        self._update_type_tab_labels()

    def _rebuild_column_type_panel(self):
        scroll_frame_to_top(self.type_scroll)
        for w in self.type_list_container.winfo_children():
            w.destroy()
        self.col_kind_vars.clear()
        self.col_type_rows.clear()
        self.kind_combos.clear()
        if self.df is None:
            self._update_type_tab_labels()
            return

        self.type_tab_var.set("all")
        self._highlight_type_tab("all")
        self.auto_col_kinds = detect_all_column_kinds(self.df)
        for col in self.df.columns:
            auto_kind = self.auto_col_kinds[col]
            auto_label = COLUMN_KIND_LABELS[auto_kind]
            var = tk.StringVar(value=auto_label)
            self.col_kind_vars[col] = var

            row = ctk.CTkFrame(self.type_list_container, fg_color=CARD_BG, corner_radius=4)
            self.col_type_rows[col] = row
            ctk.CTkLabel(
                row, text=str(col), anchor="w", font=FONT_SMALL, text_color=TEXT_PRIMARY
            ).pack(side="left", fill="x", expand=True, padx=8, pady=4)
            ctk.CTkLabel(
                row, text=auto_label, width=64, font=FONT_SMALL, text_color=TEXT_SECONDARY
            ).pack(side="left")
            combo = ctk.CTkComboBox(
                row,
                variable=var,
                values=KIND_OPTIONS,
                width=88,
                font=FONT_SMALL,
                command=lambda _v, al=auto_label, v=var, cb=None, c=col: self._on_kind_change(al, v, c),
            )
            combo.pack(side="left", padx=4, pady=4)
            self.kind_combos[col] = combo

        self._refresh_type_tab_view()
        self._scroll_panels_to_top()

    def _on_kind_change(self, auto_label: str, var: tk.StringVar, col: str) -> None:
        combo = self.kind_combos.get(col)
        if combo:
            if var.get() != auto_label:
                combo.configure(border_color=DANGER)
            else:
                combo.configure(border_color=BORDER)
        self._refresh_type_tab_view()
        self._rebuild_filters_for_column(col)

    def _rebuild_filters_for_column(self, col: str) -> None:
        if self.filter_panel and col in self.filter_panel.editors:
            self.filter_panel.editors[col]._rebuild_cond_widgets()

    def _reset_column_kinds(self):
        for col, var in self.col_kind_vars.items():
            auto_kind = self.auto_col_kinds.get(col, "text")
            var.set(COLUMN_KIND_LABELS[auto_kind])
            combo = self.kind_combos.get(col)
            if combo:
                combo.configure(border_color=BORDER)
        self._refresh_type_tab_view()

    def _add_split_column(self):
        col = self.split_col_var.get().strip()
        if not col:
            _show_warning(self.root, "提示", "请先选择要添加的拆分列")
            return
        if self.df is None:
            _show_warning(self.root, "提示", "请先加载 Excel 文件")
            return
        self.split_panel.add(col)
        self._add_filter(col=col)

    def _remove_split_column(self):
        removed = self.split_panel.remove_selected()
        if removed:
            self._remove_filter_for_column(removed)

    def _clear_filters(self) -> None:
        if self.filter_panel:
            self.filter_panel.clear()

    def _has_filter_for_column(self, col: str) -> bool:
        return bool(self.filter_panel and col in self.filter_panel.editors)

    def _remove_filter_for_column(self, col: str) -> None:
        if self.filter_panel and col in self.filter_panel._items:
            self.filter_panel.remove_col(col)

    def _add_filter(self, col: str | None = None) -> None:
        if self.df is None or not self.filter_panel:
            _show_warning(self.root, "提示", "请先加载 Excel 文件")
            return
        col = (col or self.filter_col_var.get()).strip()
        if not col:
            _show_warning(self.root, "提示", "请先选择要筛选的列")
            return
        try:
            self.filter_panel.add(self.df, col)
        except Exception as e:
            _show_error(self.root, "错误", f"添加筛选失败:\n{e}")

    def _remove_filter(self) -> None:
        if self.filter_panel:
            self.filter_panel.remove_selected()

    def _refresh_output_hint(self) -> None:
        if self.output_mode_var.get() == OUTPUT_MODE_WORKBOOK:
            self._set_out_hint(
                f"{{源文件名}}_{OUTPUT_FOLDER_LABEL}_{{时间戳}}.xlsx  （拆分后自动生成）"
            )
        else:
            self._set_out_hint(
                f"{{源文件名}}_{OUTPUT_FOLDER_LABEL}_{{时间戳}}/  （拆分后自动生成文件夹）"
            )

    def _open_path(self, path: str | Path) -> None:
        path = str(Path(path))
        try:
            if sys.platform == "win32":
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.run(["open", path], check=False)
            else:
                subprocess.run(["xdg-open", path], check=False)
        except Exception as e:
            _show_error(self.root, "错误", f"无法打开:\n{e}")

    def _show_complete_dialog(self, message: str, result: SplitExportResult) -> None:
        win = ctk.CTkToplevel(self.root)
        win.title("完成")
        win.transient(self.root)
        win.grab_set()
        win.resizable(False, False)
        apply_window_icon(win)

        top_bar = ctk.CTkFrame(win, fg_color=SUCCESS, height=4, corner_radius=0)
        top_bar.pack(fill="x")

        body = ctk.CTkFrame(win, fg_color=CARD_BG)
        body.pack(fill="both", expand=True)

        ctk.CTkLabel(
            body,
            text=message,
            font=FONT_BODY,
            text_color=TEXT_PRIMARY,
            justify="left",
            wraplength=440,
        ).pack(anchor="w", padx=24, pady=(20, 16))

        btn_row = ctk.CTkFrame(body, fg_color="transparent")
        btn_row.pack(fill="x", padx=24, pady=(0, 20))

        if result.mode == OUTPUT_MODE_WORKBOOK:
            primary_button(
                btn_row,
                text="打开文件",
                command=lambda: self._open_path(result.output_path),
                width=100,
            ).pack(side="left", padx=(0, 8))
            secondary_button(
                btn_row,
                text="打开所在文件夹",
                command=lambda: self._open_path(result.output_path.parent),
                width=120,
            ).pack(side="left", padx=(0, 8))
        else:
            primary_button(
                btn_row,
                text="打开文件夹",
                command=lambda: self._open_path(result.output_path),
                width=100,
            ).pack(side="left", padx=(0, 8))
        secondary_button(btn_row, text="确定", command=win.destroy, width=80).pack(side="left")

        win.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - win.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - win.winfo_height()) // 2
        win.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def _run_split(self):
        if self.df is None or not self.file_path:
            _show_warning(self.root, "提示", "请先选择并加载 Excel 文件")
            return

        split_cols = self.split_panel.get_all()
        if not split_cols:
            _show_warning(self.root, "提示", "请至少添加一列用于拆分")
            return

        filters = []
        for ed in (self.filter_panel.editors.values() if self.filter_panel else []):
            f = ed.to_filter_dict()
            if f:
                filters.append(f)

        try:
            self.run_btn.configure(state="disabled")
            self._show_progress(True)
            self.progress_bar.set(0)
            self._set_status("正在处理…")
            self.root.update_idletasks()
            data = self.df.copy()
            original_len = len(data)
            if filters:
                data = apply_filters(data, filters)
            if filters and len(data) == original_len and any(
                f.get("kind") == "date" for f in filters
            ):
                _show_warning(
                    self.root,
                    "筛选结果异常",
                    f"日期筛选后仍为全部 {original_len} 行，筛选可能未生效。\n"
                    "请检查列类型是否为「日期」，或重新选择筛选列刷新控件。",
                )
                return
            if data.empty:
                hint = ""
                for f in filters:
                    if f.get("kind") == "date" and f["column"] in self.df.columns:
                        s = pd.to_datetime(self.df[f["column"]], errors="coerce")
                        if s.notna().any():
                            hint = (
                                f"\n\n「{f['column']}」数据年份约在 "
                                f"{int(s.dt.year.min())}–{int(s.dt.year.max())} 年，请核对筛选范围。"
                            )
                        break
                _show_info(self.root, "结果", f"筛选后没有数据，请调整筛选条件。{hint}")
                self._set_status("筛选结果为空")
                return

            mode = self.output_mode_var.get()
            if mode == OUTPUT_MODE_WORKBOOK:
                out_path = make_timestamped_output_file(self.file_path)
            else:
                out_path = make_timestamped_output_dir(self.file_path)
            col_kinds = self._get_col_kinds_dict()

            def on_progress(current: int, total: int) -> None:
                self._report_split_progress(current, total)

            result = split_and_export(
                data,
                split_cols,
                out_path,
                header_rows=self.header_rows,
                col_kinds=col_kinds,
                mode=mode,
                on_progress=on_progress,
            )
            self.progress_bar.set(1)
            self._set_out_hint(str(result.output_path))
            if mode == OUTPUT_MODE_WORKBOOK:
                unit = "个工作表"
                detail = f"按 {', '.join(split_cols)} 拆分为 {result.count} 个工作表"
            else:
                unit = "个文件"
                detail = f"按 {', '.join(split_cols)} 拆分为 {result.count} 个文件"
            self._set_status(f"完成：生成 {result.count} {unit}")
            self._show_complete_dialog(
                f"筛选后共 {len(data)} 行\n"
                f"{detail}\n\n"
                f"保存位置:\n{result.output_path}",
                result,
            )
        except Exception as e:
            _show_error(self.root, "错误", str(e))
            self._set_status("处理失败")
        finally:
            self.run_btn.configure(state="normal")
            self._show_progress(False)

    def run(self):
        self.root.mainloop()


def main():
    from ui_theme import set_windows_app_user_model_id

    set_windows_app_user_model_id()
    ExcelSplitApp().run()


if __name__ == "__main__":
    main()
