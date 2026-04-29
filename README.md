# urlCheck

**Author:** Jamal Mazrui
**License:** MIT

`urlCheck` is a Windows tool that checks web pages for accessibility problems. It opens each page in Microsoft Edge, runs the [axe-core](https://github.com/dequelabs/axe-core) testing engine, and saves a set of output files in a new folder named after the page title.

You can use it from the command line or from a small parameter dialog. The dialog is designed to be friendly to screen readers.

The whole urlCheck project may be downloaded as a single zip archive from:

<https://github.com/JamalMazrui/urlCheck/archive/main.zip>

---

## What you need

- Windows 10 or later (64-bit)
- Microsoft Edge (any modern version; already present on Windows 10/11)
- An internet connection during each scan

You do **not** need to install Python or .NET separately. The installer ships everything urlCheck needs, and the .NET Framework 4.8 used by the parameter dialog ships in-box with Windows 10 (since version 1903) and Windows 11.

---

## Installing

Run `urlCheck_setup.exe` (the setup wizard) and follow the prompts. By default urlCheck installs to `C:\Program Files\urlCheck`, adds a Start-menu shortcut, and adds a desktop shortcut whose hotkey is **Alt+Ctrl+U**. Pressing Alt+Ctrl+U from anywhere in Windows opens the urlCheck dialog.

The installer never adds a right-click verb or any other Explorer integration. It does not change file associations.

---

## Running urlCheck

There are two ways to run it.

### From the dialog (easiest)

Launch urlCheck from any of these:

- The desktop shortcut (or its **Alt+Ctrl+U** hotkey)
- The Start-menu shortcut
- Double-clicking `urlCheck.exe` in File Explorer
- A Run dialog (`Win+R`) typing `urlCheck`

The parameter dialog appears. Fill in the fields you want and press OK to start the scan. Press F1 inside the dialog for in-context help.

The dialog has these controls. Each label has an underlined letter that you can press with **Alt** to jump straight to that control (the underlined letter is shown in brackets below):

- **Source URLs** [S] — what to scan. Enter one URL, or a domain name like `microsoft.com`, or several URLs separated by spaces, or the path to a plain text file that lists URLs (and/or local file paths) one per line. The Browse source [B] button opens a file picker.
- **Output directory** [O] — where the per-scan folders go. Blank means the current working directory. The Choose output [C] button opens a folder picker.
- **Invisible mode** [I] — run Edge with no visible browser window. Off by default.
- **View output** [V] — when all scans complete, open the output directory in File Explorer.
- **Log session** [L] — write a fresh `urlCheck.log` file in the current working directory. Useful for diagnostics; replaces any prior log.
- **Use configuration** [U] — load these field values from a saved configuration file at startup, and save them back when you press OK. The configuration lives at `%LOCALAPPDATA%\urlCheck\urlCheck.ini`. Without this checkbox, urlCheck leaves no settings on disk.
- **Help** [H] — show this help summary and offer to open the full README in your browser. F1 also shows Help.
- **Default settings** [D] — clear all fields, uncheck all boxes, and delete the saved configuration if any.
- **OK** / **Cancel** — start the scan, or cancel without scanning. Enter is OK; Esc is Cancel.

When all scans complete, a final results dialog summarizes which URLs were processed and what their page titles were.

### From the command line

`urlCheck.cmd` is a tiny wrapper script that suppresses a harmless internal warning. Use it instead of `urlCheck.exe` directly.

Show help:

```cmd
urlCheck -h
```

Show version:

```cmd
urlCheck -v
```

Check a single page:

```cmd
urlCheck https://example.com
```

Check a page by domain (`https://` is added automatically):

```cmd
urlCheck microsoft.com
```

Check several pages in one run:

```cmd
urlCheck https://a.example.com https://b.example.com microsoft.com
```

Check a list of URLs from a file:

```cmd
urlCheck urls.txt
```

The dialog can also be launched from the command line with `-g`:

```cmd
urlCheck -g
```

When invoked without arguments from a GUI shell (Explorer double-click, Start-menu shortcut, desktop hotkey), urlCheck shows the dialog automatically. When invoked without arguments from a console shell (cmd.exe, PowerShell, Windows Terminal), it prints help and exits. The `-g` flag forces GUI mode regardless.

---

## Scanning a list of URLs

If the source is a path to an existing file, urlCheck treats each non-blank, non-comment line as one target. The file may have any extension; urlCheck looks at its contents to confirm it is plain text. Inside the file:

- A line that looks like a URL (`https://...`) or a domain (`example.com`) is fetched in Edge.
- A line that is a path to a local file is loaded as HTML in Edge regardless of the file's extension. The user is responsible for choosing files that are HTML-renderable.
- Blank lines are ignored.
- Lines starting with `#` are treated as comments and ignored.

Example list file `urls.txt`:

```
# Pages to check - April 2026
https://example.com
https://example.org/about
microsoft.com
C:\work\demo.html
D:\drafts\newsletter.htm
```

Each target gets its own output folder. The accessibility report is **not** opened automatically after each scan; the user opens the files when ready.

---

## Output files

Each scan creates a new folder named after the page title (sanitized for the file system) under the output directory. If you don't choose an output directory, urlCheck uses the current working directory. From a console (cmd.exe, PowerShell), that is wherever you ran the command. From the installed shortcuts (Start menu, desktop, Alt+Ctrl+U hotkey), it is your Documents folder (`%USERPROFILE%\Documents`).

Inside each per-scan folder:

### report.htm

The main accessibility report, in HTML, with headings and links. This is the file most users will read first. It groups violations by rule, shows the failing element's HTML and CSS selector, and links to the axe-core documentation for each rule.

### report.csv

A spreadsheet of violations, one row per failing element. All fields are quoted so the file opens cleanly in Excel even when HTML attributes contain quotes. Columns include the timestamp, page title, page URL, browser version, axe-core source, outcome, rule id, impact, description, help text, help URL, tags, WCAG references, standards references, target selector, HTML snippet, and failure summary.

### report.xlsx

An Excel workbook with four sheets:

- **Metadata** — URL, page title, browser version, scan time, file names
- **Summary** — counts of violations by impact level and most common rules
- **Results** — the same data as `report.csv`, one row per issue
- **Glossary** — definitions of terms used in the report

### results.json

The complete raw output from the scan, plus metadata. Includes all four outcome types: violations, incomplete results that need manual review, passes, and inapplicable rules.

### page.yaml

The ARIA accessibility tree of the page, in YAML form. Shows the hierarchical structure of accessible elements with their roles, accessible names, ARIA attributes (level, checked, expanded, etc.), and text content. Only nodes visible to assistive technologies are included. UTF-8 with a byte-order mark.

### page.htm

A snapshot of the page as the browser saw it after rendering. External CSS and JavaScript files are inlined where the program could retrieve them.

### page.png

A full-page screenshot taken at a 1600 x 1440 viewport.

---

## Configuration file

When **Use configuration** is checked in the dialog (or `-u` is on the command line), urlCheck reads and writes a small INI file at:

```
%LOCALAPPDATA%\urlCheck\urlCheck.ini
```

It stores the source value, the output directory, and the option checkboxes. Without **Use configuration**, urlCheck leaves nothing on disk between runs. **Default settings** in the dialog deletes this file.

---

## Log file

When **Log session** is checked (or `-l` is on the command line), urlCheck writes a fresh `urlCheck.log` to the current working directory. Any prior log is deleted at the start of the run, so the file always reflects only the current session.

The current working directory depends on how urlCheck was launched. From a console (cmd.exe, PowerShell), it is whatever directory the console is in. From the installed shortcuts (Start menu, desktop, Alt+Ctrl+U hotkey), it is your Documents folder (`%USERPROFILE%\Documents`). The shortcuts deliberately point at Documents rather than the install directory so that output and the log file land somewhere writable.

The log captures: program version, Python version, architecture (32 or 64-bit), platform, frozen-exe status, working directory, command-line arguments, GUI auto-detection inputs and decision, .NET runtime info (in GUI mode), thread apartment state changes, dialog ShowDialog enter/exit, file and folder picker enter/exit, axe-core fetch results, per-URL navigation, output directory paths, and per-scan summary lines.

The log is UTF-8 with a byte-order mark, so Notepad opens it correctly.

---

## Notes

- urlCheck only reports automatically-detectable violations. It does not replace manual testing or screen-reader testing.
- urlCheck pre-fetches axe-core once per run from `cdn.jsdelivr.net`, falling back to `unpkg.com` if the first is unreachable. If both CDNs are unreachable, the scan will fail.
- Local files inside a URL list are loaded as HTML regardless of extension. urlCheck does not validate file contents before loading; if a file is not HTML, Edge may render it unexpectedly. The user is responsible for choosing HTML-renderable files.
- urlCheck never opens output files automatically. After a scan, the user can open `report.htm` (or any other output file) by visiting the per-scan folder. Use **View output** (or `--view-output`) to have the parent output directory opened in Explorer at the end of a run.

---

## Development

This section is for developers who want to build `urlCheck.exe` from source or modify it.

### Source layout

The whole program is one Python file: `urlCheck.py`. It uses [Playwright](https://playwright.dev/python/) to drive Microsoft Edge, [axe-core](https://github.com/dequelabs/axe-core) for the accessibility scan, [openpyxl](https://openpyxl.readthedocs.io/) for the Excel workbook, and [pythonnet](https://pythonnet.github.io/) for the WinForms parameter dialog. The build is packaged with [PyInstaller](https://pyinstaller.org/) into a single self-contained 64-bit exe. The installer is built with [Inno Setup](https://jrsoftware.org/isinfo.php) from `urlCheck.iss`.

### Prerequisites

- Python 3.13 (recommended; pythonnet's officially supported maximum at time of writing). If you use a newer Python that pythonnet does not yet have wheels for, the build will fail at the `pip install pythonnet` step.
- Inno Setup 6.x to compile the installer from `urlCheck.iss`.
- An internet connection during the first build (to install Python packages and the Edge Playwright driver).

### Coding style

The source uses what the author calls "Camel Type": Hungarian prefix notation for variables (`b` for boolean, `i` for integer, `s` for string, `l` for list, `d` for dict, `o` for other object types, etc.), lower camelCase for everything, single-line `if-then` statements where appropriate, and constants prefixed with `c_`. Function-level variable definitions appear at the top of each function in alphabetical order, one type per line. The style is tuned for screen-reader productivity (predictable token shapes that read well aloud).

### Building the executable

Run the included script. It checks for Python 3.13, installs the required packages, downloads the Playwright Edge driver, and runs PyInstaller:

```cmd
buildUrlCheck.cmd
```

The result is `dist\urlCheck.exe` plus the wrapper script `dist\urlCheck.cmd` and the icon `dist\urlCheck.ico`.

To build manually:

```cmd
py -3.13 -m pip install --upgrade pip
py -3.13 -m pip install --upgrade -r requirements.txt
py -3.13 -m pip install "pyinstaller>=6.19.0"
py -3.13 -m playwright install msedge
py -3.13 -m PyInstaller --clean --noconfirm --onefile --console ^
    --name urlCheck --icon=urlCheck.ico ^
    --collect-all pythonnet ^
    --hidden-import playwright.sync_api ^
    urlCheck.py
```

The `--collect-all pythonnet` flag is required for the WinForms dialog to load inside the frozen exe; without it, pythonnet's `Python.Runtime.dll` bridge is missing at runtime.

### Building the installer

After `buildUrlCheck.cmd` produces `dist\urlCheck.exe`, copy the rest of the source tree into `dist\` so that Inno Setup can find them:

```
dist\urlCheck.exe        (built; the icon is embedded inside it)
dist\urlCheck.cmd        (copied by buildUrlCheck.cmd)
dist\urlCheck.ico        (copied by buildUrlCheck.cmd; needed only at
                          installer compile time for the setup wizard's
                          own icon, not shipped with the installed program)
dist\README.htm          (you provide; rendered from README.md)
dist\README.md           (you provide)
dist\license.htm         (you provide)
dist\announce.htm        (you provide; release notes)
dist\announce.md         (you provide)
dist\urlCheck.py         (you provide; the source)
dist\buildUrlCheck.cmd   (you provide)
dist\urlCheck.iss        (you provide)
dist\requirements.txt    (you provide)
```

Open `urlCheck.iss` in Inno Setup and click Compile. The result is `dist\urlCheck_setup.exe`.

The runtime distribution from the installer's perspective is minimal: `urlCheck.exe` (with embedded icon) plus `urlCheck.cmd`. Everything else in the install directory (README, license, source, etc.) is documentation and reference material.

The installer:

- Installs to `C:\Program Files\urlCheck` by default
- Adds a desktop shortcut whose hotkey is Alt+Ctrl+U
- Adds a Start-menu group with shortcuts to urlCheck, the README, and the uninstaller
- Adds no right-click Explorer verbs and no file associations
- Is 64-bit only

### Running from source

To run urlCheck without building the executable:

```cmd
py -3.13 urlCheck.py microsoft.com
py -3.13 urlCheck.py -g
```

You will see a small bootloader warning at exit when running from source (a known issue with how PyInstaller handles `MSVCP140.dll`). The frozen `urlCheck.exe` plus `urlCheck.cmd` wrapper hides this warning; running from source you can ignore it.

### Notes on architecture

- urlCheck is built as a 64-bit console-subsystem executable. Microsoft Edge on Windows 10/11 is 64-bit, so architecture parity avoids subtle issues; pythonnet's wheels are best-tested for 64-bit Python.
- Console-subsystem means the exe gets a console when launched from Explorer. urlCheck detects GUI launches and hides its console window in that case via `ShowWindow(SW_HIDE)`. The detection uses `GetConsoleProcessList` (the same approach 2htm uses) with parent-process inspection as a fallback if the count is ambiguous.
- The WinForms dialog runs on the calling thread set to the single-threaded apartment (STA) state. This is required for common dialogs (`OpenFileDialog`, the underlying COM-shell dialogs).
- The `OpenFileDialog` (Browse source) sets `AutoUpgradeEnabled = False` to use the legacy Win32 GetOpenFileName common dialog rather than the modern Vista-era `IFileOpenDialog`. The modern dialog deadlocks under pythonnet (issues #657 and #1286 in the pythonnet repo, both unresolved at time of writing).
- The Choose output button bypasses WinForms entirely and calls `SHBrowseForFolderW` from `shell32.dll` directly via `ctypes`. It uses the classic dialog style (no `BIF_NEWDIALOGSTYLE`) because the new dialog style requires `OleInitialize` and a guaranteed STA apartment, which pythonnet doesn't reliably provide.

### Uninstalling

Use Apps & Features in Windows Settings, or run the uninstaller from the urlCheck Start-menu group. The uninstaller removes the program files. It does not touch `%LOCALAPPDATA%\urlCheck\urlCheck.ini` or any `urlCheck.log` files in working directories — delete those manually if you want a fully clean removal.

---

## License

MIT License. See `license.htm`.
