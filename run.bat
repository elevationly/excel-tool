@echo off
chcp 65001 >nul
cd /d "%~dp0"
REM 固定 Python 3.8，目标运行环境：Windows 7
py -3.8 -c "from ui_theme import set_windows_app_user_model_id; set_windows_app_user_model_id(); import excel_split_gui; excel_split_gui.main()"
