@echo off
rem urlCheck.cmd -- run urlCheck from the command line, from the top of the
rem installed folder or of the project. The program lives in exec\, as in
rem every Homer app; arguments pass straight through.
rem
rem 2>nul hides one cosmetic line PyInstaller's bootloader can print when it
rem cannot remove its temporary folder at exit. urlCheck writes its own
rem messages to standard output, so nothing of the program's is lost, and its
rem session log in %LOCALAPPDATA%\urlCheck\logs has everything.
echo Launching Edge
"%~dp0exec\urlCheck.exe" %* 2>nul
