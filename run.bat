@echo off
cd /d "%~dp0"
python -c "from ui_theme import set_windows_app_user_model_id; set_windows_app_user_model_id(); import excel_split_gui; excel_split_gui.main()"
