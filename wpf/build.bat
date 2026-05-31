@echo off
chcp 65001 >nul
cd /d "%~dp0"

where dotnet >nul 2>&1
if errorlevel 1 goto no_dotnet

echo Step 1 - restore...
dotnet restore ExcelSplitTool.sln
if errorlevel 1 goto fail

echo Step 2 - publish Ant Design Blazor 桌面版 (net8.0-windows)...
taskkill /F /IM "Excel表格拆分工具.exe" >nul 2>&1
if exist dist rmdir /s /q dist
dotnet publish ExcelSplitTool.Blazor\ExcelSplitTool.Blazor.csproj -c Release -o dist
if errorlevel 1 goto fail

if not exist "dist\Excel表格拆分工具.exe" (
    echo ERROR: exe not found in dist
    goto fail
)

echo Step 3 - trim unused Ant Design theme CSS (~5MB)...
for %%F in (
    ant-design-blazor.dark.css ant-design-blazor.dark.min.css
    ant-design-blazor.variable.css ant-design-blazor.variable.min.css
    ant-design-blazor.aliyun.css ant-design-blazor.aliyun.min.css
    ant-design-blazor.compact.css ant-design-blazor.compact.min.css
) do if exist "dist\wwwroot\_content\AntDesign\css\%%F" del /q "dist\wwwroot\_content\AntDesign\css\%%F"

if /i "%~1"=="with-redist" goto copy_redist
if /i "%~2"=="with-redist" goto copy_redist
goto build_done

:copy_redist
echo Step 4 - copy Win7 redist into dist...
if exist "..\win7_安装运行库.bat" copy /Y "..\win7_安装运行库.bat" "dist\" >nul
if exist "..\redist\win7" (
    if not exist "dist\redist\win7" mkdir "dist\redist\win7"
    xcopy /E /I /Y "..\redist\win7" "dist\redist\win7\" >nul
)
if exist "..\redist\内网部署必读.txt" copy /Y "..\redist\内网部署必读.txt" "dist\" >nul

:build_done
echo.
echo ========================================
echo   BUILD OK - Ant Design Blazor UI
echo   Output: dist\Excel表格拆分工具.exe
echo   Win10+ needs .NET 8 desktop runtime
echo   Add redist: wpf\build.bat with-redist
echo   Win7 package: root 打包_Win7内网部署包.bat
echo ========================================
echo.
if /i not "%~1"=="nopause" pause
exit /b 0

:no_dotnet
echo ERROR: dotnet not found. Install .NET SDK 8.
pause
exit /b 1

:fail
echo BUILD FAILED.
pause
exit /b 1
