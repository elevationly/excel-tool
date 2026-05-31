@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion
cd /d "%~dp0"

title 打包 Win7 内网部署包

echo ============================================================
echo   Excel 表格拆分工具 — Win7 内网离线部署包
echo   （外网打一次包，U 盘整目录拷入内网，无需联网）
echo ============================================================
echo.

set "PKG=release\Win7内网部署包"
set "MISSING=0"

:: ---------- WPF 文件夹发布 或 Python 单 exe ----------
set "WPF_DIST=wpf\dist"
set "PY_EXE=dist\Excel表格拆分工具.exe"
set "USE_WPF=0"
if exist "%WPF_DIST%\Excel表格拆分工具.exe" set "USE_WPF=1"

if not exist "%WPF_DIST%\Excel表格拆分工具.exe" if not exist "%PY_EXE%" (
    echo [提示] 未找到已编译程序，将先编译 WPF 版...
    echo.
    call "%~dp0wpf\build.bat" nopause
    if errorlevel 1 (
        echo.
        echo 也可先运行根目录 build.bat 生成 Python 版，再重新执行本脚本。
        pause
        exit /b 1
    )
    if exist "%WPF_DIST%\Excel表格拆分工具.exe" set "USE_WPF=1"
)

if not exist "%WPF_DIST%\Excel表格拆分工具.exe" if not exist "%PY_EXE%" (
    echo [错误] 仍未找到可部署的程序
    pause
    exit /b 1
)
if exist "%WPF_DIST%\Excel表格拆分工具.exe" set "USE_WPF=1"

echo [1/4] 创建部署目录 %PKG% ...
if exist "%PKG%" rmdir /s /q "%PKG%"
mkdir "%PKG%" 2>nul

echo [2/4] 复制主程序 ...
if "!USE_WPF!"=="1" (
    echo        WPF 版：复制整个 wpf\dist 目录（exe + dll）
    xcopy /E /I /Y "%WPF_DIST%\*" "%PKG%\" >nul
) else (
    echo        Python 版：单文件 exe
    copy /Y "%PY_EXE%" "%PKG%\" >nul
)

echo [3/4] 复制 Win7 离线安装脚本与说明 ...
copy /Y "win7_安装运行库.bat" "%PKG%\" >nul
if exist "redist\win7\离线安装包说明.txt" (
    copy /Y "redist\win7\离线安装包说明.txt" "%PKG%\redist\win7\" >nul
)
if exist "redist\win7\文件清单.txt" (
    copy /Y "redist\win7\文件清单.txt" "%PKG%\redist\win7\" >nul
)
if exist "redist\win7\ndp48\离线安装说明.txt" (
    if not exist "%PKG%\redist\win7\ndp48" mkdir "%PKG%\redist\win7\ndp48"
    copy /Y "redist\win7\ndp48\离线安装说明.txt" "%PKG%\redist\win7\ndp48\" >nul
)

echo [4/4] 复制离线运行库（x64 / x86）...
if exist "redist\win7\x64" (
    xcopy /E /I /Y "redist\win7\x64" "%PKG%\redist\win7\x64\" >nul
) else (
    echo        [缺少] redist\win7\x64\  — 64 位机无法离线安装
    set "MISSING=1"
)
if exist "redist\win7\x86" (
    xcopy /E /I /Y "redist\win7\x86" "%PKG%\redist\win7\x86\" >nul
) else (
    echo        [缺少] redist\win7\x86\  — 32 位机无法离线安装
    set "MISSING=1"
)
if exist "redist\win7\ndp48\ndp48-x86-x64-allos-enu.exe" (
    if not exist "%PKG%\redist\win7\ndp48" mkdir "%PKG%\redist\win7\ndp48"
    copy /Y "redist\win7\ndp48\ndp48-x86-x64-allos-enu.exe" "%PKG%\redist\win7\ndp48\" >nul
)

:: 写入简短内网说明
> "%PKG%\内网部署必读.txt" (
    echo Excel 表格拆分工具 — Win7 内网部署
    echo.
    echo 【拷贝】将本文件夹整包拷入内网 Win7 电脑（保持目录结构不变）。
    echo.
    echo 【安装运行库】右键「win7_安装运行库.bat」→ 以管理员身份运行 → 输入 Y
    echo   - 仅需运行一次；不访问互联网
    echo   - 需 Windows 7 SP1
    echo.
    echo 【WPF 版】若使用 wpf 打包的 exe，还需 .NET Framework 4.8：
    echo   - 将 ndp48-x86-x64-allos-enu.exe 放入 redist\win7\ndp48\ 后重新打离线包
    echo   - 或内网已统一安装过 .NET 4.8 则可直接运行 exe
    echo.
    echo 【运行】双击 Excel表格拆分工具.exe
    echo   WPF 版：勿单独拷 exe，须保留同目录全部 dll 与 .config
    echo.
    echo 详细说明见 redist\win7\离线安装包说明.txt
)

echo.
echo ============================================================
if !MISSING! equ 1 (
    echo   打包完成，但离线运行库文件不完整！
    echo.
    echo   请在外网电脑按 redist\win7\离线安装包说明.txt 下载：
    echo     x64\ 与 x86\ 各 3 个文件（.msu x2 + vc_redist.exe）
    echo   放入本仓库 redist\win7\ 后重新运行本脚本。
) else (
    echo   打包完成：%PKG%
    echo.
    echo   内网步骤：
    echo     1. 整包拷入 Win7
    echo     2. 管理员运行 win7_安装运行库.bat
    echo     3. 运行 Excel表格拆分工具.exe
)
echo ============================================================
echo.
pause
exit /b 0
