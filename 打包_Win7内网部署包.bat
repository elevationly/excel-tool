@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion
cd /d "%~dp0"

title Win7 offline package

echo ============================================================
echo   Excel split tool - Win7 offline package
echo ============================================================
echo.

set "PKG=release\Win7内网部署包"
set "MISSING=0"

set "WPF_DIST=wpf\dist"
set "PY_EXE=dist\Excel表格拆分工具.exe"
set "USE_WPF=0"
if exist "%WPF_DIST%\Excel表格拆分工具.exe" set "USE_WPF=1"

if not exist "%WPF_DIST%\Excel表格拆分工具.exe" if not exist "%PY_EXE%" (
    echo [hint] Building wpf\dist ...
    echo.
    call "%~dp0wpf\build.bat" with-redist nopause
    if errorlevel 1 (
        echo.
        echo Or run root build.bat for Python exe, then run this script again.
        pause
        exit /b 1
    )
    if exist "%WPF_DIST%\Excel表格拆分工具.exe" set "USE_WPF=1"
)

if not exist "%WPF_DIST%\Excel表格拆分工具.exe" if not exist "%PY_EXE%" (
    echo [ERROR] No built exe found.
    pause
    exit /b 1
)
if exist "%WPF_DIST%\Excel表格拆分工具.exe" set "USE_WPF=1"

echo [1/4] Create %PKG% ...
if exist "%PKG%" rmdir /s /q "%PKG%"
mkdir "%PKG%" 2>nul

echo [2/4] Copy program ...
if "!USE_WPF!"=="1" (
    xcopy /E /I /Y "%WPF_DIST%\*" "%PKG%\" >nul
) else (
    copy /Y "%PY_EXE%" "%PKG%\" >nul
)

echo [3/4] Copy Win7 install scripts ...
copy /Y "win7_安装运行库.bat" "%PKG%\" >nul
if not exist "%PKG%\redist\win7" mkdir "%PKG%\redist\win7"
if exist "redist\win7\离线安装包说明.txt" copy /Y "redist\win7\离线安装包说明.txt" "%PKG%\redist\win7\" >nul
if exist "redist\win7\文件清单.txt" copy /Y "redist\win7\文件清单.txt" "%PKG%\redist\win7\" >nul
if not exist "%PKG%\redist\win7\ndp48" mkdir "%PKG%\redist\win7\ndp48"
if exist "redist\win7\ndp48\离线安装说明.txt" copy /Y "redist\win7\ndp48\离线安装说明.txt" "%PKG%\redist\win7\ndp48\" >nul
if exist "redist\内网部署必读.txt" copy /Y "redist\内网部署必读.txt" "%PKG%\内网部署必读.txt" >nul

echo [4/4] Copy offline redist x64/x86 ...
if exist "redist\win7\x64" (
    xcopy /E /I /Y "redist\win7\x64" "%PKG%\redist\win7\x64\" >nul
) else (
    echo        [MISSING] redist\win7\x64
    set "MISSING=1"
)
if exist "redist\win7\x86" (
    xcopy /E /I /Y "redist\win7\x86" "%PKG%\redist\win7\x86\" >nul
) else (
    echo        [MISSING] redist\win7\x86
    set "MISSING=1"
)
if exist "redist\win7\ndp48\ndp48-x86-x64-allos-enu.exe" (
    copy /Y "redist\win7\ndp48\ndp48-x86-x64-allos-enu.exe" "%PKG%\redist\win7\ndp48\" >nul
)

echo.
echo ============================================================
if !MISSING! equ 1 (
    echo   DONE with warnings: redist files incomplete.
    echo   See redist\win7\offline readme txt for download list.
) else (
    echo   DONE: %PKG%
    echo   Steps: copy folder to Win7 -^> run win7_安装运行库.bat as admin -^> run exe
)
echo ============================================================
echo.
pause
exit /b 0
