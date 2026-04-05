@echo off
setlocal
urlCheck.exe %* 2>&1 | findstr /V "Failed to remove temporary directory"
endlocal
