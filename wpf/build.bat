@echo off
chcp 65001 >nul
cd /d "%~dp0"

where dotnet >nul 2>&1
if errorlevel 1 goto no_dotnet

echo Step 1 - restore NuGet...
dotnet restore ExcelSplitTool.sln
if errorlevel 1 goto fail

echo Step 2 - publish Release (exe + dll 同目录)...
taskkill /F /IM "Excel表格拆分工具.exe" >nul 2>&1
if exist dist rmdir /s /q dist
dotnet publish ExcelSplitTool\ExcelSplitTool.csproj -c Release -o dist
if errorlevel 1 goto fail

if not exist "dist\Excel表格拆分工具.exe" (
    echo ERROR: exe not found in dist
    goto fail
)

echo Step 3 - copy Win7 offline redist...
if exist "..\win7_安装运行库.bat" copy /Y "..\win7_安装运行库.bat" "dist\" >nul
if exist "..\redist\win7" xcopy /E /I /Y "..\redist\win7" "dist\redist\win7\" >nul

echo.
echo ========================================
echo   BUILD OK
echo   部署: 将整个 dist 文件夹拷到目标电脑
echo   运行: dist\Excel表格拆分工具.exe
echo   Win7: 先管理员运行 dist\win7_安装运行库.bat
echo   完整内网包: 仓库根目录 打包_Win7内网部署包.bat
echo ========================================
echo.
if /i not "%~1"=="nopause" pause
exit /b 0

:no_dotnet
echo.
echo ERROR: dotnet command not found.
echo Install .NET SDK 8 from https://dotnet.microsoft.com/download
echo Or: winget install Microsoft.DotNet.SDK.8
echo.
pause
exit /b 1

:fail
echo.
echo BUILD FAILED. See messages above.
echo.
pause
exit /b 1
