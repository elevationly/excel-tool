@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo 使用 Python 3.8 安装依赖（Win7 兼容）...
py -3.8 -m pip install -r requirements.txt -q

echo.
echo 正在打包为单个 exe（Python 3.8 / Win7，约需 1~3 分钟）...
py -3.8 -m PyInstaller --noconfirm --clean ^
  --onefile ^
  --windowed ^
  --name "Excel表格拆分工具" ^
  --icon "assets\app_icon.ico" ^
  --add-data "assets\app_icon.ico;assets" ^
  --collect-all customtkinter ^
  --collect-all tkcalendar ^
  --collect-all babel ^
  --hidden-import "tkcalendar" ^
  --hidden-import "tkcalendar.calendar_" ^
  --hidden-import "tkcalendar.dateentry" ^
  --hidden-import "babel" ^
  --hidden-import "babel.dates" ^
  --hidden-import "openpyxl.cell._writer" ^
  excel_split_gui.py

if errorlevel 1 (
    echo.
    echo 打包失败，请检查上方错误信息。
    pause
    exit /b 1
)

echo.
echo 完成！输出文件：
echo   dist\Excel表格拆分工具.exe
echo.
pause
