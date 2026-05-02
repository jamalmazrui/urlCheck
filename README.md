---
title: "urlCheck — Accessibility Checker for Web Pages"
author: "Jamal Mazrui"
description: "Accessibility Checker for Web Pages"
---

# urlCheck

**Author:** Jamal Mazrui
**License:** MIT

`urlCheck` is one of three companion accessibility tools by Jamal Mazrui:

- **2htm** — convert documents (Word, Excel, PowerPoint, PDF, Markdown) to accessible HTML
- **extCheck** — check Office and Markdown files for accessibility problems
- **urlCheck** — check web pages for accessibility problems

The three tools share a common command-line and GUI layout, so learning one makes the others easy to pick up.

`urlCheck` is a Windows tool that checks web pages for accessibility problems. It opens each page in Microsoft Edge, runs the [axe-core](https://github.com/dequelabs/axe-core) testing engine, and saves a set of output files in a new folder named after the page title.

Like its companion tools, `urlCheck` runs in two modes: a **GUI mode** (a small parameter dialog launched by double-clicking the program, pressing its desktop hotkey, or running with `-g`) and a **command-line mode** (any other invocation, suitable for batch files and pipelines). Both modes accept the same options.

---

## What you need

- Windows 10 or later (64-bit)
- Microsoft Edge (already present on Windows 10/11; urlCheck uses your installed Edge directly and does not bundle or download a separate browser)
- An internet connection during each scan

You do **not** need to install Python or .NET separately. The installer ships everything `urlCheck` needs, and the .NET Framework 4.8 used by the parameter dialog ships in-box with Windows 10 (since version 1903) and Windows 11.

---

## Installing

Download `urlCheck_setup.exe` from the [GitHub repository](https://github.com/JamalMazrui/urlCheck) and run it. The setup wizard:

- Prompts you for the installation directory (default: `C:\Program Files\urlCheck`).
- Includes a brief MIT license summary on the welcome page; the full license text is installed alongside the program as `License.htm`.
- Adds a Start-menu shortcut and a desktop shortcut whose hotkey is **Alt+Ctrl+U**. Pressing **Alt+Ctrl+U** from anywhere in Windows opens the `urlCheck` dialog.

The final wizard page offers two checkboxes (both checked by default): launch `urlCheck` (with a hotkey reminder) and read the HTML documentation.

---

## Running urlCheck

### From the dialog (easiest)

Launch `urlCheck` from any of these:

- The desktop shortcut (or its **Alt+Ctrl+U** hotkey)
- The Start-menu shortcut
- Double-clicking `urlCheck.exe` in File Explorer
- A Run dialog (`Win+R`) typing `urlCheck`

The parameter dialog has these controls. Each label has an underlined letter that you can press with **Alt** to jump straight to that control:

- **Source urls** [S] — one url (https://example.com), or a domain (microsoft.com), or several of either separated by spaces, or the path to a single plain text file that lists urls, domains, or local HTML file paths one per line. The list file may have any extension; urlCheck verifies it is plain text by inspecting its contents.
- **Browse source...** [B] — pick a single source from a file picker
- **Output directory** [O] — where the output is written. Blank means the current working directory.
- **Choose output...** [C] — pick the output directory from a folder picker
- **Invisible mode** [I] — run Edge with no visible browser window
- **Authenticate credentials** [A] — pause after each newly-encountered domain so you can sign in / dismiss cookie banners / accept popups, then press Enter on the console to resume the scan. Forces a visible browser. CLI-only at present.
- **Force replacements** [F] — reuse an existing per-page output folder by emptying its contents and writing a fresh set of files. Without this, urlCheck skips the url when its per-page output folder already exists, preserving previous scan results.
- **View output** [V] — open the output directory in File Explorer when the run is done
- **Log session** [L] — write a fresh `urlCheck.log` in the output directory (or current directory if no output directory is set)
- **Use configuration** [U] — load these field values from the saved configuration at startup, and save them back when you press OK
- **Help** [H] — show this help summary and offer to open the full README. F1 also shows Help.
- **Default settings** [D] — clear all fields, uncheck all boxes, and delete the saved configuration if any
- **OK** / **Cancel** — start the run, or cancel without running. Enter is OK; Esc is Cancel.

The Browse source and Choose output pickers open at the directory derived from the corresponding text field's current value when that value points to an existing path; otherwise they open at your Documents folder. With **Use configuration** checked, those text fields are pre-populated from your last session, so the pickers naturally pick up where you left off.

If you press OK with an output directory that does not yet exist, urlCheck prompts to create it (default Yes). Choosing No keeps the dialog open with focus on the output field so you can correct it.

When all pages have been processed, a final results dialog summarizes what was done.


### From the command line

Open a Command Prompt and run `urlCheck` with the source as an argument:

```cmd
# Single URL:
urlCheck https://example.com

# Several URLs:
urlCheck https://a.com https://b.com

# URLs from a file:
urlCheck urls.txt

# Output to a directory:
urlCheck *.htm -o reports

# View output when done:
urlCheck https://example.com --view-output

# Open the GUI:
urlCheck -g

```

When invoked without arguments from a GUI shell (Explorer double-click, Start-menu shortcut, desktop hotkey), `urlCheck` shows the dialog automatically. When invoked without arguments from a console shell, it prints help and exits. The `-g` flag forces GUI mode regardless.

---

## Command-line options

| Option | Long form | Description |
|---|---|---|
| `-h` | `--help` | Show usage and exit |
| `-v` | `--version` | Show version and exit |
| `-g` | `--gui-mode` | Show the parameter dialog |
| `-o <d>` | `--output-dir <d>` | Write output to `<d>` (created if missing); defaults to current directory |
| `-f` | `--force` | reuse an existing per-page output folder by emptying its contents and writing a fresh set of files |
|   | `--view-output` | After the run, open the output directory in File Explorer |
| `-l` | `--log` | Write `urlCheck.log` (UTF-8 with BOM) in the output directory; replaced each session |
| `-u` | `--use-configuration` | Read saved defaults from `%LOCALAPPDATA%\urlCheck\urlCheck.ini` |
| `-i` | `--invisible` | Run Microsoft Edge with no visible browser window |
| `-a` | `--authenticate` | Pause on first url of each new hostname for the user to authenticate, then press Enter to continue |

Every option in the GUI corresponds one-to-one with a command-line flag, so a workflow prototyped in the dialog can be translated to a batch file without surprises.

---

## Supported sources

urlCheck accepts:

- A single url (`https://example.com` or just `example.com`)
- Several urls separated by spaces
- The path to a plain text file with one url or local HTML file path per line; the file may have any extension

Url-list files are detected automatically by content sniffing, not by extension.

---

## Output

For each scanned page, urlCheck creates a folder named after the page title and writes these files inside it:

- `report.htm` — human-readable accessibility report (open in any browser)
- `report.csv` — the same findings in tabular form
- `report.xlsx` — Excel workbook with separate sheets for violations, passes, incomplete, and inapplicable rules
- `results.json` — full structured scan output (metadata + axe-core results) for programmatic downstream use
- `page.yaml` — ARIA accessibility tree of the page
- `page.htm` — saved page source with stylesheet hrefs preserved
- `page.png` — full-page screenshot

If a per-page folder with the same sanitized title already exists, urlCheck skips that url by default — previous scan results are preserved. Use `--force` (or check **Force replacements** in the dialog) to instead empty the existing folder and replace its contents with a fresh scan. The skip decision is made right after the page title is read, before the expensive accessibility scan, so re-running urlCheck on a long url list is cheap when most pages have already been scanned.

If `--view-output` is set, the **parent** output directory (the one containing the per-page subfolders) opens in File Explorer at the end of the run.

---

## Configuration file

When **Use configuration** is checked in the dialog (or `-u` is on the command line), `urlCheck` reads and writes a small INI file at:

```
%LOCALAPPDATA%\urlCheck\urlCheck.ini
```

It stores the source field, the output directory, and the option checkboxes. Without **Use configuration**, `urlCheck` leaves nothing on disk between runs. **Default settings** in the dialog deletes this file.

---

## Log file

When **Log session** is checked (or `-l` is on the command line), `urlCheck` writes a fresh `urlCheck.log` to the output directory (or current directory if no output directory is set). Any prior log is replaced at the start of the run, so the file always reflects only the current session.

The log captures: program version, command-line arguments, GUI auto-detection, the resolved output directory, per-page events, and any errors (including tracebacks for unexpected failures).

Without **Log session**, `urlCheck` does not create any log or error file on disk. Errors are reported only to the console (and the GUI results dialog, in GUI mode).

The log is UTF-8 with a byte-order mark, so Notepad opens it correctly.

---

## Notes

- urlCheck reports the violations the [axe-core](https://github.com/dequelabs/axe-core) engine detects automatically. It does not replace manual testing.
- Local files inside a URL list are loaded as HTML regardless of extension. urlCheck does not validate file contents before loading; if a file is not HTML, Edge may render it unexpectedly. The user is responsible for choosing HTML-renderable files.
- urlCheck waits for the page to finish loading and pauses briefly so late DOM updates are more likely to settle before the scan runs. Pages with very long-running asynchronous content may need a manual retry.

---

## Development

This section is for developers who want to build the executable from source. End users can skip it.

### Distribution layout

The runtime distribution shipped by `urlCheck_setup.exe` is just a few files: `urlCheck.exe` plus the HTML documentation (`ReadMe.htm`, `Announce.htm`, `License.htm`). The Markdown sources, the build script, the installer script, the icon, the program source, and the coding-style guide live in the GitHub repository (and in this `urlCheck.zip` archive).

### Source layout

The whole program is one Python file: `urlCheck.py`. It uses [Playwright](https://playwright.dev/python/) to drive Microsoft Edge, [axe-core](https://github.com/dequelabs/axe-core) for the accessibility scan, [openpyxl](https://openpyxl.readthedocs.io/) for the Excel workbook, and [pythonnet](https://pythonnet.github.io/) for the WinForms parameter dialog. The build is packaged with [PyInstaller](https://pyinstaller.org/) into a single self-contained 64-bit exe. The installer is built with [Inno Setup](https://jrsoftware.org/isinfo.php) from `urlCheck.iss`.

### Coding style

The source uses what the author calls "Camel Type" (Python variant): Hungarian prefix notation for variables (`b` for boolean, `i` for integer, `s` for string, `l` for list, `d` for dict, etc.), lower camelCase for everything where Python conventionally uses snake_case, and the lowercase class name as the prefix for non-basic-type variables. The `o` prefix is reserved for COM objects only; urlCheck does not use COM, so it has no `o`-prefixed variables. Constants follow the same Hungarian-prefixed naming as variables — only definition placement and convention conveys constant-ness. See `CamelType_Python.md` in this archive for the full guidelines.

### Edge runtime

urlCheck drives the system-installed Microsoft Edge through Playwright's `channel="msedge"` mechanism. **It does not bundle Edge** and does not run `playwright install msedge` (which would conflict with the existing system Edge on every modern Windows machine). Edge ships in-box with Windows 10 and 11; if a user has somehow removed it, urlCheck surfaces a friendly message pointing at <https://www.microsoft.com/edge>.

### Prerequisites

- Python 3.13 64-bit (recommended; pythonnet's officially supported maximum at time of writing). The build script verifies the selected Python is 64-bit before proceeding.
- Inno Setup 6.x to compile the installer from `urlCheck.iss`.
- An internet connection during the first build (to install Python packages from PyPI).

### Building the executable

Run the included script:

```cmd
buildUrlCheck.cmd
```

It auto-detects the compiler, verifies the build environment, embeds the icon into `urlCheck.exe`, and produces the runtime distribution in `dist\`.

### Building the installer

Open `urlCheck.iss` in Inno Setup and click Compile. The result is `dist\urlCheck_setup.exe`.

The installer ships only the runtime files: `urlCheck.exe` plus the HTML documentation (`ReadMe.htm`, `Announce.htm`, `License.htm`), plus the cosmetic-warning suppression wrapper (`urlCheck.cmd`). Markdown sources, the build script, this `.iss` script, the icon, the source file, and any coding-style guideline files live in the GitHub repository.

### Uninstalling

Use Apps & Features in Windows Settings, or run the uninstaller from the `urlCheck` Start-menu group. The uninstaller removes the program files. It does not touch `%LOCALAPPDATA%\urlCheck\urlCheck.ini` or any `urlCheck.log` files in working directories — delete those manually if you want a fully clean removal.


## License

MIT License. See `License.htm` (installed alongside the program) or `License.md` (in the GitHub repository).