@echo off
rem ===================================================================
rem Build urlCheck.exe from urlCheck.py.
rem
rem Requires:
rem   - Windows 10 or later (x64). The .NET Framework 4.8 ships in-box
rem     with Windows 10 (since version 1903) and Windows 11; no
rem     separate runtime install is needed.
rem   - Python 3.13 or later from https://www.python.org/downloads/
rem     (pythonnet 3.0.5 supports Python 3.7-3.13; if a newer Python
rem     becomes available before pythonnet adds support, use 3.13)
rem   - Internet access (to install Python packages and the Edge
rem     Playwright browser binary on first build)
rem
rem Outputs:
rem   - dist\urlCheck.exe (single-file executable)
rem   - dist\urlCheck.cmd (small console-warning-suppressing wrapper)
rem ===================================================================

setlocal
cd /d "%~dp0"

rem ---- Pick a Python launcher ---------------------------------------
rem
rem Prefer a 3.13 launcher (pythonnet's officially supported maximum at
rem time of writing). Fall back to whatever py picks by default. Note
rem the goto-based control flow rather than a parenthesized if-block:
rem cmd.exe's parser pre-expands variables in a parenthesized block
rem and mis-parses periods inside echoed strings, both of which can
rem produce baffling "was unexpected at this time" errors.
rem -------------------------------------------------------------------
set "c_sPy=py -3.13"
%c_sPy% --version >nul 2>&1
if not errorlevel 1 goto :pyFound

set "c_sPy=py"
%c_sPy% --version >nul 2>&1
if not errorlevel 1 goto :pyFoundFallback

echo [ERROR] No Python launcher (py) found
echo         Install Python 3.13 from https://www.python.org/downloads/
exit /b 2

:pyFoundFallback
echo [WARN] Python 3.13 launcher not found; using default py launcher

:pyFound
echo [INFO] Using launcher: %c_sPy%
%c_sPy% --version

echo Installing Python dependencies (Playwright 1.58+ for aria_snapshot,
echo pythonnet 3.0.5+ for the WinForms GUI, openpyxl for Excel output)...
%c_sPy% -m pip install --upgrade pip
if errorlevel 1 exit /b %errorlevel%

%c_sPy% -m pip install --upgrade -r requirements.txt
if errorlevel 1 exit /b %errorlevel%

%c_sPy% -m pip install "pyinstaller>=6.19.0"
if errorlevel 1 exit /b %errorlevel%

echo Installing Playwright browser support for Microsoft Edge...
%c_sPy% -m playwright install msedge
if errorlevel 1 exit /b %errorlevel%

echo Building urlCheck.exe (onefile, console subsystem)...
rem
rem PyInstaller flags:
rem
rem   --onefile             Single self-contained urlCheck.exe.
rem   --console             Console-subsystem binary so CLI output goes to
rem                         the parent shell. The program detects GUI
rem                         launches via GetConsoleProcessList and hides
rem                         its own console window in that case, so a
rem                         double-click from Explorer does not leave a
rem                         stray console behind.
rem   --icon=urlCheck.ico   Embeds the multi-resolution icon into the EXE
rem                         (16, 24, 32, 48, 64, 128, 256 sizes). Windows
rem                         picks the right resolution for Explorer, the
rem                         taskbar, the Alt+Tab switcher, the title bar,
rem                         and the Apps & Features list.
rem   --collect-all pythonnet
rem                         Bundle every pythonnet submodule and the
rem                         Python.Runtime.dll bridge so `import clr`
rem                         resolves at runtime inside the frozen exe.
rem   --hidden-import       Cover Playwright's lazy submodules.
rem     playwright.sync_api
rem
%c_sPy% -m PyInstaller --clean --noconfirm --onefile --console ^
    --name urlCheck ^
    --icon=urlCheck.ico ^
    --collect-all pythonnet ^
    --hidden-import playwright.sync_api ^
    urlCheck.py
if errorlevel 1 exit /b %errorlevel%

echo Copying wrapper script and icon into dist\ ...
copy /Y urlCheck.cmd dist\urlCheck.cmd
if errorlevel 1 exit /b %errorlevel%
rem urlCheck.ico is copied into dist\ only because Inno Setup needs the
rem file alongside urlCheck.iss at compile time (for SetupIconFile=).
rem At runtime the icon is already embedded in urlCheck.exe (via the
rem PyInstaller --icon flag above), so the .ico file does NOT need to
rem be installed alongside the exe.
copy /Y urlCheck.ico dist\urlCheck.ico
if errorlevel 1 exit /b %errorlevel%

echo(
echo Build complete. The runtime distribution is just two files:
echo   dist\urlCheck.exe   (the icon is embedded)
echo   dist\urlCheck.cmd
echo(
echo To produce the installer (urlCheck_setup.exe):
echo   1. Copy README.htm, license.htm, README.md, announce.md,
echo      announce.htm, urlCheck.py, urlCheck.ico, buildUrlCheck.cmd,
echo      urlCheck.iss, and requirements.txt into dist\
echo      (urlCheck.ico is needed here at compile time only, not at
echo      runtime, because Inno Setup uses it for the wizard's icon.)
echo   2. Open urlCheck.iss in Inno Setup and click Compile.
echo      The result is dist\urlCheck_setup.exe.
endlocal
