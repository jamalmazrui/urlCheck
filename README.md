---
title: "urlCheck — Accessibility Checker for Web Pages"
author: "Jamal Mazrui"
description: "Accessibility Checker for Web Pages"
---

# urlCheck

**Author:** Jamal Mazrui
**License:** MIT
**Project home:** <https://github.com/JamalMazrui/urlCheck>

`urlCheck` is free, open-source software released under the [MIT License](https://opensource.org/licenses/MIT). Anyone is welcome to download and use it, study its source code, and adapt it. The only requirement is a modern version of Windows; nothing else needs to be installed.

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

- Prompts you for the installation folder (default: `C:\Program Files\urlCheck`).
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
- **Output folder** [O] — where the output is written. Blank means the current working folder.
- **Choose output...** [C] — pick the output folder from a folder picker
- **Invisible mode** [I] — run Edge with no visible browser window
- **Authenticate credentials** [A] — pause after each newly-encountered domain so you can sign in / dismiss cookie banners / accept popups, then press Enter (or click OK in GUI mode) to resume the scan. By default, urlCheck uses a fresh temporary Edge profile and disconnects its automation channel from Edge during the user-interaction pause (which improves the chance of success against sites that detect active automation, such as WhatsApp Web); to use your real profile instead, also check **Main profile**. If both **Invisible mode** and **Authenticate credentials** are checked, urlCheck overrides Invisible mode at run time and launches Edge with a visible window (an auth prompt requires a visible browser); the override is logged.
- **Main profile** [M] — launch Edge with your real (default) Edge user profile so saved logins, cookies, and session state are available. Without it, urlCheck uses a fresh temporary profile so the scan is anonymous and your real profile is not exposed to the scanned site. Independent of **Authenticate credentials**. Requires that no Microsoft Edge process is already running, since Edge cannot share a profile across two processes. urlCheck checks at startup; if Edge is running, the CLI exits with a friendly message and the GUI shows a dialog explaining why it cannot proceed and asks you to close Edge before submitting again.
- **Force replacements** [F] — reuse an existing per-page output folder by emptying its contents and writing a fresh set of files. Without this, urlCheck skips the url when its per-page output folder already exists, preserving previous scan results.
- **View output** [V] — open the output folder in File Explorer when the run is done
- **Log session** [L] — write a fresh `urlCheck.log` in the output folder (or current folder if no output folder is set)
- **Use configuration** [U] — load these field values from the saved configuration at startup, and save them back when you press OK
- **Help** [H] — show this help summary and offer to open the full README. F1 also shows Help.
- **Default settings** [D] — clear all fields, uncheck all boxes, and delete the saved configuration if any
- **OK** / **Cancel** — start the run, or cancel without running. Enter is OK; Esc is Cancel.

**Note on profiles and privacy.** urlCheck's default — a fresh temporary profile — matches the experience of an anonymous member of the public visiting a site for the first time. The scan captures and analyzes whatever a brand-new visitor would see. Choosing **Main profile** is a deliberate departure from that: pages may be personalized to your account (recommendations tailored to your history, content visible only to you, your name and avatar in the header), and the captured page.htm and screenshot may include that personal information. If you plan to share or publish the output, review it before doing so.

The Browse source and Choose output pickers open at the folder derived from the corresponding text field's current value when that value points to an existing path; otherwise they open at your Documents folder. With **Use configuration** checked, those text fields are pre-populated from your last session, so the pickers naturally pick up where you left off.

If you press OK with an output folder that does not yet exist, urlCheck prompts to create it (default Yes). Choosing No keeps the dialog open with focus on the output field so you can correct it.

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

# Output to a folder:
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
| `-o <d>` | `--output-folder <d>` | Write output to `<d>` (created if missing); defaults to current folder |
| `-f` | `--force` | reuse an existing per-page output folder by emptying its contents and writing a fresh set of files |
|   | `--view-output` | After the run, open the output folder in File Explorer |
| `-l` | `--log` | Write `urlCheck.log` (UTF-8 with BOM) in the output folder. Appends across runs by default; combine with `-f` (`--force`) to replace the prior log instead. |
| `-u` | `--use-configuration` | Read saved defaults from `%LOCALAPPDATA%\urlCheck\urlCheck.ini` |
| `-i` | `--invisible` | Run Microsoft Edge with no visible browser window |
| `-a` | `--authenticate` | Pause on first url of each registrable domain for the user to authenticate, then press Enter (or click OK) to resume. By default uses a fresh temporary profile and disconnects Playwright during the pause; combine with `-m` to use your real profile (no disconnect). Auto-disables `-i`. |
| `-m` | `--main-profile` | Launch Edge with your real (default) Edge profile so saved logins are available. Without `-m`, urlCheck uses a fresh temporary profile so the scan is anonymous. Requires that no Edge process is already running. |

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

For each scanned page, urlCheck creates a subfolder whose name is based on the page title, adjusted as needed for the file system. Inside the folder:

- `report.htm` — human-readable accessibility report (open in any browser)
- `report.xlsx` — Excel workbook with separate sheets for violations, passes, incomplete, and inapplicable rules. The Results sheet has an **Image** column whose cells are clickable hyperlinks to per-violation screenshots (see below).
- `results.json` — full structured scan output (metadata + axe-core results) for programmatic downstream use
- `page.yaml` — ARIA accessibility tree of the page
- `page.htm` — saved page source with stylesheet hrefs preserved
- `page.png` — full-page screenshot
- `violations/` — element-level screenshots, one PNG per violation node where the screenshot could be captured (see below)

### Per-violation screenshots

For each rule-violation node found by axe, urlCheck attempts to capture an element-level screenshot using the CSS selector axe provides. Successful captures are saved as `violations/image-001.png`, `violations/image-002.png`, etc. The Image column of the Results sheet in `report.xlsx` shows the basename as a clickable hyperlink to the relative path; clicking the cell opens the PNG in the OS default image viewer.

The hyperlinks are **relative** (e.g., `violations/image-001.png`, not absolute paths), so you can move, zip, or share the page subfolder freely — the links resolve correctly wherever the workbook ends up, as long as the `violations/` subfolder travels with it.

When axe's CSS selector cannot be resolved by the browser engine — typical cases include shadow-DOM-pierced selectors, hidden or zero-size elements, or selectors that match multiple elements — urlCheck silently skips that node. The corresponding row of the Results sheet has an empty Image cell. There is no warning printed; the assumption is that the user looking at a Results sheet sees what was captured and what wasn't, and can use the CSS selector in the **Path** column to find the element manually if needed.

If a per-page folder with the same sanitized title already exists, urlCheck skips that url by default — previous scan results are preserved. Use `--force` (or check **Force replacements** in the dialog) to instead empty the existing folder and replace its contents with a fresh scan. The skip decision is made right after the page title is read, before the expensive accessibility scan, so re-running urlCheck on a long url list is cheap when most pages have already been scanned.

If `--view-output` is set, the **parent** output folder (the one containing the per-page subfolders) opens in File Explorer at the end of the run.

### Accessibility Conformance Report (ACR.xlsx and ACR.docx)

At the end of every run, urlCheck writes `ACR.xlsx` and `ACR.docx` in the parent output folder. Together they form a draft Accessibility Conformance Report that aggregates axe-core results across pages and maps them to WCAG 2.2 success criteria using standard VPAT 2.5 terminology (Supports, Partially Supports, Does Not Support, Not Applicable, Not Evaluated).

The first sheet of `ACR.xlsx`, **Conformance Report**, has one row per WCAG 2.2 criterion (all 86, including Level A, AA, and AAA — the obsolete 4.1.1 Parsing is omitted). The columns are:

- **Criterion** — Criterion number, name, and level, e.g., `1.1.1 Non-text Content (A)`. Hyperlinked to the W3C WCAG 2.2 Quick Reference for that criterion.
- **Summary** — A single-sentence description of the criterion's intent.
- **Conformance** — The derived VPAT 2.5 conformance term. The cell is multi-line: line 1 is the verdict (Supports / Partially Supports / Does Not Support / Not Evaluated); subsequent lines list the page sheet names whose per-page Calc was `fail` or `partial` under "Not supported:", and pages whose per-page Calc was `manual` (incomplete) under "Not evaluated:". Only line 1 is visible by default; expand the row to see the page lists.
- **Manual** — Numbered manual-test steps for the criterion.
- **Result** — User-editable. Where the human reviewer records the final ACR verdict.
- **Remarks** — User-editable. The first line is auto-generated with axe-context instance counts: `Axe: fail N, pass N, incomplete N, inapplicable N`. The user can append additional remarks below.

The remaining per-page sheets are diagnostic views, one per scanned page, named after each page's subfolder. Each per-page sheet shows the criterion, summary, page-specific Calc verdict (pass / fail / partial / manual / na / unknown), and four columns (Fail, Pass, Incomplete, Inapplicable) listing the axe rule IDs that produced each outcome on that page. Each rule is shown with its instance count (e.g., `image-alt 3` means three failing image elements).

The final **Glossary** sheet defines all the terms used (axe outcome categories, urlCheck Calc values, VPAT 2.5 conformance terms, WCAG principles and levels) and links to relevant external resources.

The companion `ACR.docx` is a narrative summary with sections for Overview, Pages Analyzed, Conformance Summary, Criteria Requiring Attention, Methodology, and Resources. It is informed by report.htm's information architecture but adapted to focus on per-criterion conformance rather than per-rule diagnostics. Both files are draft assets generated by automation; the user is expected to manually verify and refine them, especially the Result and Remarks columns and any criteria marked Not Evaluated, before publishing the final ACR.

#### Calc formula table

The Calc column on per-page sheets uses these formulas:

| Calc value | Formula |
|---|---|
| **partial** | At least one rule instance fails AND at least one passes for the same criterion. |
| **fail** | At least one rule instance fails (no pass on the same criterion). |
| **manual** | At least one incomplete result (no fail, no pass). |
| **pass** | At least one rule instance passes (no fail, no incomplete). |
| **na** | All rule instances are inapplicable to the page. |
| **unknown** | No axe rules apply to this criterion. Default before scan. |

The Conformance column on the rollup sheet is derived from the per-page Calcs across all included pages (worst-result-wins semantics):

| Combined Calc | Conformance |
|---|---|
| All `pass` | Supports |
| Any `pass` and any `fail` | Partially Supports |
| Any `fail` (no pass) | Does Not Support |
| Any `manual` (no fail or pass) | Not Evaluated |
| Only `na` everywhere | Supports (vacuously satisfied per WCAG 2.0 Understanding Conformance) |
| `unknown` everywhere | Not Evaluated |

Counts are by node instance (one DOM element flagged), summed across all included pages. So if 3 image elements lack alt text, axe-core reports 3 instances of the `image-alt` violation, and the Remarks line shows `fail 3` for the matching criterion.

The "Not Evaluated" term is reserved by VPAT 2.5 for AAA criteria, but urlCheck uses it more broadly for cases where automated testing cannot reach a verdict. This is defensible because the urlCheck output is a **draft** ACR. The user is expected to perform the manual checks listed in the Manual column, write a final verdict in the Result column, and save the curated workbook in a separate folder before publishing.

#### Workbook scope

Without `--force`, urlCheck scopes the report to **all** subfolders under the parent that contain a `results.json` — every page that has ever been scanned into this output folder. To exclude a page from the report, simply delete its subfolder. Your edits to the Result and Remarks columns on the rollup sheet are preserved across re-runs (matched by criterion id).

With `--force`, urlCheck scopes the report to only the URLs scanned in the current session, ignoring older subfolders. This is the way to start a fresh ACR.

If no pages have been scanned yet (or all subfolders are excluded), `ACR.xlsx` is still written: it contains the criterion list with manual-test instructions and the Glossary, ready for the user to fill in.

#### Accessibility failure rate

Each `report.htm` per-page report and the `ACR.docx` narrative include an *accessibility failure rate* — a single number that summarizes how well a page (or a page set) is doing on automated accessibility checks. It's defined as:

```
rate = 100000 * impactWeightedInstances / pageBytes
```

where the numerator is `1*minor + 2*moderate + 3*serious + 4*critical` summed over every violation instance on the page, and the denominator is the byte size of the saved page source (`page.htm`). Each impact level reflects axe-core's own severity rating; instances are individual flagged DOM elements, not distinct rules.

The constant `100000` is tuned so the result reads naturally as a percent. **Lower is better.** A clean page is well under 1%; a typical problematic page lands in the double digits; a truly broken page can exceed 100%. The percent framing is purely a display convention to make the number easy to grasp and remember; the underlying quantity is impact-weighted violation instances per byte of page source, which has no natural ceiling.

For a page set (the ACR-level rate), urlCheck sums per-page numerators and per-page denominators before dividing. This is a size-weighted view: bigger pages contribute proportionally more, reflecting that they have more content with more places where users might encounter violations. The aggregate rate is shown in `ACR.docx`'s metadata header; per-page rates are shown next to each page in the Pages Analyzed list and in `report.xlsx`'s Summary sheet.

The metric is meant to be tracked over time. As the page owner remediates issues, the rate should drop from one scan session to the next. Accessibility is a journey, not a destination.

---

## Configuration file

When **Use configuration** is checked in the dialog (or `-u` is on the command line), `urlCheck` reads and writes a small INI file at:

```
%LOCALAPPDATA%\urlCheck\urlCheck.ini
```

It stores the source field, the output folder, and the option checkboxes. Without **Use configuration**, `urlCheck` leaves nothing on disk between runs. **Default settings** in the dialog deletes this file.

---

## Log file

When **Log session** is checked (or `-l` is on the command line), `urlCheck` writes `urlCheck.log` (UTF-8 with BOM) to the output folder (or current folder if no output folder is set). The log is opened in **append** mode by default, so accumulated history across runs is preserved — useful for diagnosing intermittent issues. To start fresh each session, also enable **Force replacements** (or pass `-f` / `--force` on the command line); the prior log is then deleted before the new one is opened. Sessions are visually separated by a blank line in append mode.

The log captures: program version, command-line arguments, GUI auto-detection, the resolved output folder, per-page events, and any errors (including tracebacks for unexpected failures).

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