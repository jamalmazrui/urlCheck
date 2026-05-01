# urlCheck Release Notes

## Version 1.10.0

### What's new

- **Edge browser handling.** urlCheck drives the system-installed Microsoft Edge through Playwright's `channel="msedge"` mechanism. The build script no longer runs `playwright install msedge`, which was actively harmful: on a Windows machine where Edge is already installed (the universal case), Playwright errors with "msedge is already installed on the system" and would fail the build. Edge ships in-box with Windows 10 and 11.
- **Friendly Edge error.** If the rare situation occurs where Playwright cannot find Edge (Edge removed by the user, unusual configuration), urlCheck now surfaces a friendly message pointing at https://www.microsoft.com/edge instead of a Python traceback.
- **64-bit build guard.** `buildUrlCheck.cmd` verifies the selected Python interpreter is 64-bit before invoking PyInstaller. A 32-bit Python is rejected with a clear error so the produced `urlCheck.exe` is consistently 64-bit.
- **Skip-or-overwrite output behavior, consistent with extCheck and 2htm.** When a per-page output folder with the same sanitized page title already exists, urlCheck now skips that URL by default rather than auto-suffixing a new folder. With `--force` (or the Force replacements checkbox in the dialog), the existing folder is emptied and a fresh scan is written into it. The skip decision is made right after the page title is read, before the expensive accessibility scan, so re-running urlCheck on a long URL list is cheap when most pages have already been scanned. Note: this is a behavior change from prior versions, which used auto-suffixed folder names like `My Page-001`.
- **Force replacements checkbox in the GUI dialog.** Matches the corresponding checkbox in extCheck and 2htm.
- **Picker initial directory.** The Browse source and Choose output buttons now open at the directory derived from the text-field value when that value points to an existing path (whether the user just typed it or it was loaded from a saved configuration), and at the user's Documents folder otherwise. The folder picker honors the initial path via a BFFCALLBACK that posts BFFM_SETSELECTION on BFFM_INITIALIZED — the previous version was unable to honor an initial path with the classic SHBrowseForFolder style.
- **Skipped count in the run summary.** A multi-URL run now reports successful, skipped, and error counts separately.
- **No implicit error files.** Earlier versions wrote an `error.txt` and a fallback `untitled-page` folder when a scan crashed unexpectedly. urlCheck now creates files only when the user explicitly asks for them. Errors are surfaced to the console (always) and to `urlCheck.log` when `-l` is on. If you run without `-l` and want full tracebacks for any errors, the run summary now reminds you to add `-l` next time.
- **Camel Type coding standard.** All variable names follow the project's Camel Type style. The `o` prefix is reserved for COM objects only; non-COM objects use the lowercase class name as their prefix. Examples: `oPath` → `path`, `oResponse` → `response`, `oDlg` → `dialog`, `oError` → `ex`. urlCheck does not use COM, so it has no remaining `o`-prefixed variables.
- **Cross-program naming.** Identifier names for shared concepts now match across the three companion tools (urlCheck, extCheck, 2htm). The program-name and version constants are `sProgramName` and `sProgramVersion`; the config and log filename constants are `sConfigDirName`, `sConfigFileName`, `sLogFileName`; the source and output-directory variables are `sSource` and `sOutputDir`; the GUI layout constants are all `iLayout*` (left, right, top, gap, rowGap, labelWidth, buttonWidth, buttonHeight, textHeight, formWidth). The `logger` class now has the same surface in all three programs: `open`, `close`, `info`, `warn`, `error`, `debug`.

### Technical notes

- The icon is embedded in `urlCheck.exe` at build time via PyInstaller's `--icon` flag (multi-resolution: 16, 24, 32, 48, 64, 128, 256). Shortcuts inherit the icon.
- Two ctypes Structure class definitions (`ProcessEntry32` and `BrowseInfoW`) were renamed from their old `o`-prefixed forms to follow Python's PascalCase class-naming convention.
- The build script (`buildUrlCheck.cmd`) writes `urlCheck.exe` to the current working directory rather than a `dist\` subfolder, and it no longer copies `urlCheck.cmd` or `urlCheck.ico` aside (those files already live in the working directory next to `urlCheck.iss`). PyInstaller's intermediate work files are stashed in a `build\` subfolder out of the way; the auto-generated `urlCheck.spec` is removed after a successful build. The script uses no forward `call :label`; it ships with CRLF line endings as defense-in-depth.
- The Invisible checkbox in the dialog is labeled "Invisible mode" (mnemonic on the I).

### Installer (`urlCheck_setup.exe`)

- 64-bit only.
- Prompts for the installation directory (default: `C:\Program Files\urlCheck`).
- Includes a brief MIT-license summary on the welcome page.
- Installs only HTML versions of the documentation (`ReadMe.htm`, `Announce.htm`, `License.htm`); the Markdown counterparts and source/build/installer scripts live in the GitHub repository.
- The "Launch urlCheck now" checkbox on the final page reminds the user that the desktop hotkey is Alt+Ctrl+U.
