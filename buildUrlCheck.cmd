@echo off
rem ===================================================================
rem buildUrlCheck.cmd  --  Build urlCheck.exe from urlCheck.py
rem
rem Compiles urlCheck.py with PyInstaller into a single self-contained
rem 64-bit executable, embedding urlCheck.ico. The exe is written to the
rem current directory (alongside the source, the icon, the wrapper .cmd,
rem and urlCheck.iss).
rem
rem Requires:
rem   - Windows 10 or later (x64). The .NET Framework 4.8 ships in-box
rem     with Windows 10 (since version 1903) and Windows 11; no
rem     separate runtime install is needed.
rem   - Python 3.13 from https://www.python.org/downloads/ (pythonnet's
rem     officially supported maximum at time of writing). Must be
rem     64-bit Python; the script verifies this before building.
rem   - Internet access (to install Python packages on first build)
rem
rem Microsoft Edge:
rem   urlCheck drives the system-installed Microsoft Edge through
rem   Playwright's `channel="msedge"` mechanism, NOT a Playwright-
rem   bundled Chromium. Edge ships with Windows 10/11 by default, so
rem   no separate browser download or installation is required during
rem   the build. We deliberately do NOT run `playwright install
rem   msedge` here -- on a system where Edge is already installed
rem   (the common case), that command warns or errors with "msedge is
rem   already installed on the system" and would fail this build.
rem
rem Build outputs (in the current directory):
rem   urlCheck.exe        -- single-file 64-bit executable, embedded icon
rem
rem To produce the installer setup.exe:
rem   Open urlCheck.iss in Inno Setup and click Compile.
rem   Inno Setup writes urlCheck_setup.exe to the same directory.
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
set "sPy=py -3.13"
%sPy% --version >nul 2>&1
if not errorlevel 1 goto :pyFound

set "sPy=py"
%sPy% --version >nul 2>&1
if not errorlevel 1 goto :pyFoundFallback

echo [ERROR] No Python launcher (py) found.
echo         Install Python 3.13 from https://www.python.org/downloads/
exit /b 2

:pyFoundFallback
echo [WARN] Python 3.13 launcher not found; using default py launcher.

:pyFound
echo [INFO] Using launcher: %sPy%
%sPy% --version

rem ---- Verify 64-bit Python -----------------------------------------
rem
rem PyInstaller produces an exe whose bitness matches the Python
rem interpreter that runs it. urlCheck is meant to be a 64-bit binary
rem (consistent with 2htm and extCheck, and with the typical 64-bit
rem Windows 10/11 environment). A 32-bit Python here would silently
rem produce a 32-bit urlCheck.exe, which we don't want. Bail with a
rem clear error if struct.calcsize("P") != 8.
rem -------------------------------------------------------------------
%sPy% -c "import struct,sys; sys.exit(0 if struct.calcsize('P') == 8 else 2)"
if errorlevel 1 (
    echo [ERROR] The selected Python is not 64-bit.
    echo         urlCheck must be built with 64-bit Python so the
    echo         resulting urlCheck.exe is also 64-bit. Install the
    echo         64-bit Python 3.13 from python.org and rerun.
    exit /b 2
)
echo [INFO] Python is 64-bit (good).

rem ---- Verify the icon exists ---------------------------------------
if not exist urlCheck.ico (
    echo [ERROR] urlCheck.ico not found in %CD%.
    echo         The icon file is required to embed into the exe.
    exit /b 1
)

echo [INFO] Installing Python dependencies (Playwright 1.58+ for
echo        aria_snapshot, pythonnet 3.0.5+ for the WinForms GUI,
echo        openpyxl for Excel output)...
%sPy% -m pip install --upgrade pip
if errorlevel 1 exit /b %errorlevel%

%sPy% -m pip install --upgrade -r requirements.txt
if errorlevel 1 exit /b %errorlevel%

%sPy% -m pip install "pyinstaller>=6.19.0"
if errorlevel 1 exit /b %errorlevel%

rem Note: no `playwright install msedge` step. urlCheck uses the
rem system-installed Edge via channel="msedge"; no Playwright-bundled
rem browser binaries are needed at build time. Running
rem `playwright install msedge` on a Windows machine that already has
rem Edge produces a "msedge is already installed" warning/error, so
rem omitting it makes the build more robust as well as smaller.

echo [INFO] Building urlCheck.exe (onefile, console subsystem)...
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
rem                         Required: pythonnet ships runtime resource
rem                         files that PyInstaller's static analysis
rem                         would otherwise miss.
rem   --hidden-import       Cover Playwright's lazy submodules.
rem     playwright.sync_api
rem
rem   --distpath .          Place the final exe in the current directory
rem                         instead of dist\.
rem   --workpath build      Keep the intermediate build directory out of
rem                         the way (PyInstaller default is build\).
rem
rem We deliberately do NOT pass --specpath. PyInstaller's --icon path is
rem resolved relative to the .spec file's directory at .spec generation
rem time; routing the .spec into a subfolder makes PyInstaller look for
rem urlCheck.ico in that subfolder and fail. Letting the .spec live in
rem the current directory keeps the icon lookup simple. The leftover
rem urlCheck.spec is removed after a successful build.
rem
%sPy% -m PyInstaller --clean --noconfirm --onefile --console ^
    --name urlCheck ^
    --icon=urlCheck.ico ^
    --collect-all pythonnet ^
    --hidden-import playwright.sync_api ^
    --distpath . ^
    --workpath build ^
    urlCheck.py
if errorlevel 1 exit /b %errorlevel%

rem Clean up the auto-generated spec file. PyInstaller would re-create
rem it on the next build anyway.
if exist urlCheck.spec del /q urlCheck.spec

echo(
echo [INFO] Build complete. urlCheck.exe is in %CD% (icon embedded).
echo(
echo To produce the installer (urlCheck_setup.exe):
echo   Open urlCheck.iss in Inno Setup and click Compile.
echo   Inno Setup writes urlCheck_setup.exe to %CD%.
endlocal
