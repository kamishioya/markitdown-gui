@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
pushd "%SCRIPT_DIR%"

for %%D in (build dist .pyinstaller-work release) do (
    if exist "%%D" rmdir /s /q "%%D"
)

pyinstaller pyinstaller\markitdown-gui.spec --noconfirm --clean --distpath release --workpath .pyinstaller-work
if errorlevel 1 (
    popd
    exit /b %errorlevel%
)

if exist ".pyinstaller-work" rmdir /s /q ".pyinstaller-work"

echo.
echo Build complete.
echo Launch this executable:
echo %SCRIPT_DIR%release\MarkItDownGUI\MarkItDownGUI.exe

popd
endlocal