"""Element UI 风格主题（CustomTkinter）。"""
from __future__ import annotations

import sys
from pathlib import Path

import customtkinter as ctk

APP_USER_MODEL_ID = "eleva.excel.split.tool.1.0"

# Element UI 色板
PRIMARY = "#409EFF"
PRIMARY_HOVER = "#66B1FF"
PRIMARY_ACTIVE = "#3A8EE6"
SUCCESS = "#67C23A"
WARNING = "#E6A23C"
DANGER = "#F56C6C"

BG_PAGE = "#F2F6FC"
CARD_BG = "#FFFFFF"
BORDER = "#DCDFE6"
BORDER_LIGHT = "#EBEEF5"

TEXT_PRIMARY = "#303133"
TEXT_REGULAR = "#606266"
TEXT_SECONDARY = "#909399"
TEXT_PLACEHOLDER = "#C0C4CC"

HEADER_BG = "#409EFF"
HEADER_FG = "#FFFFFF"

FONT_TITLE = ("Microsoft YaHei UI", 20, "bold")
FONT_SECTION = ("Microsoft YaHei UI", 13, "bold")
FONT_BODY = ("Microsoft YaHei UI", 12)
FONT_SMALL = ("Microsoft YaHei UI", 11)
FONT_HINT = ("Microsoft YaHei UI", 10)

ASSETS_DIR = Path(__file__).resolve().parent / "assets"
APP_ICON = ASSETS_DIR / "app_icon.ico"


def set_windows_app_user_model_id() -> None:
    """在创建窗口前调用，使任务栏使用自定义图标而非 python.exe 默认图标。"""
    if sys.platform != "win32":
        return
    try:
        import ctypes

        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
    except Exception:
        pass


def shorten_path(path: str | Path, max_len: int = 52) -> str:
    """中间省略的路径展示。"""
    p = Path(path)
    text = str(p)
    if len(text) <= max_len:
        return text
    name = p.name
    if len(name) >= max_len - 1:
        return "…" + name[-(max_len - 1) :]
    prefix_len = max_len - len(name) - 2
    parent = str(p.parent)
    if len(parent) <= prefix_len:
        return text
    sep = "\\" if "\\" in text else "/"
    return f"…{parent[-prefix_len:]}{sep}{name}"


def scroll_frame_to_top(frame: ctk.CTkScrollableFrame) -> None:
    """将 CTkScrollableFrame 滚动条复位到顶部。"""
    try:
        canvas = frame._parent_canvas
        canvas.update_idletasks()
        canvas.yview_moveto(0)
        canvas.xview_moveto(0)
        frame.update_idletasks()
    except Exception:
        pass


def configure_app() -> None:
    set_windows_app_user_model_id()
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")
    theme = ctk.ThemeManager.theme
    theme["CTkButton"]["fg_color"] = [PRIMARY, PRIMARY]
    theme["CTkButton"]["hover_color"] = [PRIMARY_HOVER, PRIMARY_HOVER]
    theme["CTkButton"]["text_color"] = ["#FFFFFF", "#FFFFFF"]
    theme["CTkButton"]["corner_radius"] = 6
    theme["CTkEntry"]["corner_radius"] = 4
    theme["CTkEntry"]["border_color"] = [BORDER, BORDER]
    theme["CTkComboBox"]["corner_radius"] = 4
    theme["CTkComboBox"]["border_color"] = [BORDER, BORDER]
    theme["CTkComboBox"]["button_color"] = [PRIMARY, PRIMARY]
    theme["CTkComboBox"]["button_hover_color"] = [PRIMARY_HOVER, PRIMARY_HOVER]
    theme["CTkFrame"]["corner_radius"] = 8
    theme["CTkScrollableFrame"]["corner_radius"] = 6
    theme["CTkRadioButton"]["fg_color"] = [PRIMARY, PRIMARY]
    theme["CTkRadioButton"]["hover_color"] = [PRIMARY_HOVER, PRIMARY_HOVER]
    theme["CTkCheckBox"]["fg_color"] = [PRIMARY, PRIMARY]
    theme["CTkSegmentedButton"]["selected_color"] = [PRIMARY, PRIMARY]
    theme["CTkSegmentedButton"]["selected_hover_color"] = [PRIMARY_HOVER, PRIMARY_HOVER]


def apply_window_icon(window) -> None:
    if not APP_ICON.exists():
        return
    icon = str(APP_ICON)
    try:
        window.iconbitmap(default=icon)
    except Exception:
        try:
            window.iconbitmap(icon)
        except Exception:
            pass
    try:
        from PIL import Image, ImageTk

        img = Image.open(APP_ICON)
        photo = ImageTk.PhotoImage(img)
        window._app_icon_ref = photo  # 防止被 GC
        window.iconphoto(True, photo)
    except Exception:
        pass


def card(parent, **kwargs) -> ctk.CTkFrame:
    opts = dict(
        fg_color=CARD_BG,
        corner_radius=8,
        border_width=1,
        border_color=BORDER,
    )
    opts.update(kwargs)
    return ctk.CTkFrame(parent, **opts)


def section_title(parent, text: str) -> ctk.CTkLabel:
    return ctk.CTkLabel(
        parent,
        text=text,
        font=FONT_SECTION,
        text_color=TEXT_PRIMARY,
        anchor="w",
    )


def hint_label(parent, text: str) -> ctk.CTkLabel:
    return ctk.CTkLabel(
        parent,
        text=text,
        font=FONT_HINT,
        text_color=TEXT_SECONDARY,
        anchor="w",
    )


def primary_button(parent, text: str, command=None, **kwargs) -> ctk.CTkButton:
    opts = dict(
        text=text,
        command=command,
        font=FONT_BODY,
        fg_color=PRIMARY,
        hover_color=PRIMARY_HOVER,
        text_color="#FFFFFF",
        corner_radius=6,
        height=34,
    )
    opts.update(kwargs)
    return ctk.CTkButton(parent, **opts)


def secondary_button(parent, text: str, command=None, **kwargs) -> ctk.CTkButton:
    opts = dict(
        text=text,
        command=command,
        font=FONT_BODY,
        fg_color=CARD_BG,
        hover_color=BORDER_LIGHT,
        text_color=TEXT_REGULAR,
        border_width=1,
        border_color=BORDER,
        corner_radius=6,
        height=32,
    )
    opts.update(kwargs)
    return ctk.CTkButton(parent, **opts)


def danger_button(parent, text: str, command=None, **kwargs) -> ctk.CTkButton:
    opts = dict(
        text=text,
        command=command,
        font=FONT_SMALL,
        fg_color=CARD_BG,
        hover_color="#FEF0F0",
        text_color=DANGER,
        border_width=1,
        border_color="#FBC4C4",
        corner_radius=6,
        width=64,
        height=28,
    )
    opts.update(kwargs)
    return ctk.CTkButton(parent, **opts)
