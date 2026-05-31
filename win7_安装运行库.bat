@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
cd /d "%~dp0"

title Win7 离线运行库安装

echo ============================================================
echo   Excel 表格拆分工具 — Windows 7 离线运行库安装
echo   （内网环境：不访问互联网，仅使用本机 redist\win7 文件）
echo ============================================================
echo.

:: ---------- 管理员权限 ----------
net session >nul 2>&1
if errorlevel 1 (
    echo [错误] 请右键本文件，选择「以管理员身份运行」。
    echo.
    pause
    exit /b 1
)

:: ---------- 系统版本提示 ----------
ver | find "6.1" >nul
if errorlevel 1 (
    echo [提示] 当前系统可能不是 Windows 7，本脚本主要针对 Win7 SP1。
    echo        若确为 Win7 可继续；否则请确认是否需要安装。
    echo.
)

:: ---------- 判断 x64 / x86 ----------
set "ARCH=x86"
if /i "%PROCESSOR_ARCHITECTURE%"=="AMD64" set "ARCH=x64"
if /i "%PROCESSOR_ARCHITEW6432%"=="AMD64" set "ARCH=x64"

:: ---------- 离线包目录（按架构分子目录）----------
set "PKGDIR=%~dp0redist\win7\%ARCH%"
if not exist "%PKGDIR%\" (
    echo [错误] 未找到目录：%PKGDIR%
    echo.
    echo   内网部署请确认已整包拷贝，且与本 bat 同级存在：
    echo     redist\win7\%ARCH%\  （三个离线文件）
    echo.
    echo   外网打包容器应使用「打包_Win7内网部署包.bat」生成完整目录。
    echo   详见 redist\win7\离线安装包说明.txt
    echo.
    pause
    exit /b 1
)

echo [信息] 检测到系统架构：%ARCH%
echo [信息] 离线包目录：%PKGDIR%
echo.

:: ---------- 查找 KB2999226 ----------
set "MSU="
for %%F in ("%PKGDIR%\Windows6.1-KB2999226-%ARCH%.msu") do set "MSU=%%~fF"
if not defined MSU (
    for %%F in ("%PKGDIR%\*KB2999226*.msu") do set "MSU=%%~fF"
)
if not defined MSU (
    echo [错误] 未找到 KB2999226 离线包（.msu）。
    echo        请将 Windows6.1-KB2999226-%ARCH%.msu 放入：
    echo        %PKGDIR%
    echo        外网下载说明见：redist\win7\离线安装包说明.txt
    echo.
    pause
    exit /b 1
)

:: ---------- 查找可选 KB3118401 ----------
set "MSU311="
for %%F in ("%PKGDIR%\Windows6.1-KB3118401-%ARCH%.msu") do set "MSU311=%%~fF"
if not defined MSU311 (
    for %%F in ("%PKGDIR%\*KB3118401*.msu") do set "MSU311=%%~fF"
)

:: ---------- 查找 VC++ 运行库 ----------
set "VC="
for %%F in ("%PKGDIR%\vc_redist.%ARCH%.exe") do set "VC=%%~fF"
if not defined VC (
    for %%F in ("%PKGDIR%\vc_redist*.exe") do set "VC=%%~fF"
)
if not defined VC (
    echo [错误] 未找到 Visual C++ 运行库离线包（vc_redist.%ARCH%.exe）。
    echo        请将 vc_redist.%ARCH%.exe 放入：
    echo        %PKGDIR%
    echo.
    pause
    exit /b 1
)

echo 将安装以下离线包：
echo   [1] %MSU%
if defined MSU311 echo   [可选] %MSU311%
echo   [2] %VC%
echo.
set /p CONFIRM=确认安装？输入 Y 后回车，其它键取消：
if /i not "!CONFIRM!"=="Y" (
    echo 已取消。
    pause
    exit /b 0
)
echo.

set "NEED_REBOOT=0"

:: ---------- 安装 KB3118401（可选，失败不中断）----------
if defined MSU311 (
    echo ---------- 安装 KB3118401（可选）----------
    wusa.exe "%MSU311%" /quiet /norestart
    set "ERR=!ERRORLEVEL!"
    if !ERR! equ 0 (
        echo [完成] KB3118401 安装成功。
        set "NEED_REBOOT=1"
    ) else if !ERR! equ 2359302 (
        echo [跳过] KB3118401 已安装或无需安装。（代码 2359302）
    ) else (
        echo [警告] KB3118401 返回代码 !ERR!，可查阅日志后继续尝试 KB2999226。
    )
    echo.
)

:: ---------- 安装 KB2999226 ----------
echo ---------- 安装 KB2999226（Universal CRT）----------
wusa.exe "%MSU%" /quiet /norestart
set "ERR=!ERRORLEVEL!"
if !ERR! equ 0 (
    echo [完成] KB2999226 安装成功。
    set "NEED_REBOOT=1"
) else if !ERR! equ 2359302 (
    echo [跳过] KB2999226 已安装。（代码 2359302）
) else (
    echo [失败] KB2999226 安装失败，代码：!ERR!
    echo.
    echo 常见原因：
    echo   - 未安装 Windows 7 SP1
    echo   - 需先安装 KB3118401 后重试
    echo   - 补丁与当前系统语言/版本不匹配
    echo.
    echo 详细说明见 redist\win7\离线安装包说明.txt
    pause
    exit /b 1
)
echo.

:: ---------- 可选：.NET Framework 4.8 离线（WPF 版需要）----------
set "NDP48=%~dp0redist\win7\ndp48\ndp48-x86-x64-allos-enu.exe"
if exist "%NDP48%" (
    echo ---------- 安装 .NET Framework 4.8（离线）----------
    reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release 2>nul | find "528040" >nul
    if not errorlevel 1 (
        echo [跳过] 已检测到 .NET Framework 4.8 或更高版本。
    ) else (
        echo [信息] 正在静默安装，约需数分钟...
        "%NDP48%" /quiet /norestart
        set "ERR=!ERRORLEVEL!"
        if !ERR! equ 0 (
            echo [完成] .NET Framework 4.8 安装成功。
            set "NEED_REBOOT=1"
        ) else if !ERR! equ 1641 (
            echo [跳过] .NET 4.8 已安装。（代码 1641）
        ) else if !ERR! equ 3010 (
            echo [完成] .NET 4.8 已安装，需重启。（代码 3010）
            set "NEED_REBOOT=1"
        ) else (
            echo [警告] .NET 4.8 安装返回代码：!ERR!
        )
    )
    echo.
) else (
    echo [提示] 未找到 .NET 4.8 离线包（WPF 版可能需要）：
    echo        %NDP48%
    echo        说明见 redist/win7/ndp48/离线安装说明.txt
    echo.
)

:: ---------- 安装 VC++ 运行库 ----------
echo ---------- 安装 Visual C++ 2015-2022 运行库 ----------
"%VC%" /install /quiet /norestart
set "ERR=!ERRORLEVEL!"
if !ERR! equ 0 (
    echo [完成] VC++ 运行库安装成功。
) else if !ERR! equ 1638 (
    echo [跳过] 已安装相同或更高版本 VC++ 运行库。（代码 1638）
) else if !ERR! equ 3010 (
    echo [完成] VC++ 运行库已安装，建议重启。（代码 3010）
    set "NEED_REBOOT=1"
) else (
    echo [警告] VC++ 安装返回代码：!ERR!（0 为成功，1638 为已存在）
)
echo.

:: ---------- 检查关键 DLL ----------
echo ---------- 检查 UCRT 组件 ----------
set "DLL_OK=0"
if exist "%SystemRoot%\System32\api-ms-win-core-sysinfo-l1-2-0.dll" set "DLL_OK=1"
if exist "%SystemRoot%\System32\downlevel\api-ms-win-core-sysinfo-l1-2-0.dll" set "DLL_OK=1"
if !DLL_OK! equ 1 (
    echo [正常] 已检测到 api-ms-win-core-sysinfo 相关组件。
) else (
    echo [提示] 尚未在 System32 检测到该 DLL。
    echo        若刚安装补丁，请重启后再运行本脚本检查，或直接试运行主程序。
)
echo.

echo ============================================================
if !NEED_REBOOT! equ 1 (
    echo   安装完成。建议重启计算机后再运行 Excel表格拆分工具.exe。
) else (
    echo   安装完成。可直接运行 Excel表格拆分工具.exe 测试是否还有弹窗。
)
echo ============================================================
echo.
pause
exit /b 0
