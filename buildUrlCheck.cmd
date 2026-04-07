@echo off
setlocal

echo Installing Python dependencies (includes Playwright 1.58+ for aria_snapshot support)...
py -3.14 -m pip install --upgrade pip
if errorlevel 1 exit /b %errorlevel%

py -3.14 -m pip install --upgrade -r requirements.txt
if errorlevel 1 exit /b %errorlevel%

py -3.14 -m pip install "pyinstaller>=6.19.0"
if errorlevel 1 exit /b %errorlevel%

echo Installing Playwright browser support for Edge...
py -3.14 -m playwright install msedge
if errorlevel 1 exit /b %errorlevel%

echo Building urlCheck.exe...
py -3.14 -m PyInstaller --clean --noconfirm --onefile --name urlCheck urlCheck.py
if errorlevel 1 exit /b %errorlevel%

echo Copying wrapper script...
copy /Y urlCheck.cmd dist\urlCheck.cmd
if errorlevel 1 exit /b %errorlevel%

echo.
echo Build complete. Distribute both files from dist\:
echo   dist\urlCheck.exe
echo   dist\urlCheck.cmd
endlocal
