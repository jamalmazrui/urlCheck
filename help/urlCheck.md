---
title: "urlCheck Guide"
author: "Jamal Mazrui"
description: "Accessibility Checker for Web Pages"
---

# urlCheck Guide

**Author:** Jamal Mazrui
**Copyright:** © 2026 Jamal Mazrui
**License:** [MIT](https://opensource.org/license/mit/)
**Project home:** <https://github.com/JamalMazrui/urlCheck>

`urlCheck` is a Windows tool that checks web pages for accessibility problems. It opens each page in Microsoft Edge, runs the [axe-core](https://github.com/dequelabs/axe-core) testing engine, and saves a set of output files in a new folder named after the page title.

This is the full guide. For a quick start, see ReadMe. Like its companion tools `2htm` and `extCheck` (see Announce for a description of the family), `urlCheck` runs in two modes: a **GUI mode** (a small parameter dialog launched by double-clicking the program, pressing its desktop hotkey, or running with `-g`) and a **command-line mode** (any other invocation, suitable for batch files and pipelines). Both modes accept the same options.

---

## What you need

- Windows 10 or later (64-bit)
- Microsoft Edge (already present on Windows 10/11; urlCheck uses your installed Edge directly and does not bundle or download a separate browser)
- An internet connection during each scan

You do **not** need to install Python or .NET separately. The installer ships everything `urlCheck` needs, and the .NET Framework 4.8 used by the parameter dialog ships in-box with Windows 10 (since version 1903) and Windows 11.

---

## Installing

Download `urlCheck_setup.exe` from the [urlCheck releases page](https://github.com/JamalMazrui/urlCheck/releases) and run it. It needs administrator rights, because it installs for everyone on the computer. The setup wizard:

- Asks for the installation folder the first time (default: `C:\Program Files\urlCheck`). An update goes to the same folder without asking.
- Shows a short MIT license summary on the welcome page. The full license is installed as `License.htm`.
- Adds a Start-menu group and a desktop shortcut whose hotkey is **Alt+Control+Shift+U**. Pressing **Alt+Control+Shift+U** from anywhere in Windows opens the `urlCheck` dialog.

The last page offers two checkboxes. **Launch urlCheck now** is checked; **Open the user guide** is not. When you press Finish, a short results box says what was installed and where the logs are. `urlCheck` starts after you close that box, so the box is never hidden behind it.

The program itself is in the `exec` folder of the installation, and this guide and the other documents are in the `help` folder. `ReadMe.htm` and `License.htm` are at the top.

### Updating

In the dialog, press **F11**. `urlCheck` asks GitHub whether a newer version exists. If one does, Yes is the default: press Enter and the new setup program is downloaded and started. If you already have the newest version, No is the default. If the check fails (for example, with no internet connection), a message says so.

---

## Running urlCheck

### From the dialog (easiest)

Launch `urlCheck` from any of these:

- The desktop shortcut (or its **Alt+Control+Shift+U** hotkey)
- The Start-menu shortcut
- Double-clicking `urlCheck.exe` in the `exec` folder in File Explorer

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
- **Log session** [L] — also write `urlCheck.log` in the output folder (or current folder if no output folder is set). A session log is always kept in `%LOCALAPPDATA%\urlCheck\logs` either way.
- **Use configuration** [U] — load these field values from the saved configuration at startup, and save them back when you press OK
- **Guide** [G] — open this guide in your browser, then return to the dialog with everything as it was.
- **Help** [H] — list every field with its tip, the dialog's keys, and the version check. F1 also shows Help.
- **Default settings** [D] — clear all fields, uncheck all boxes, and delete the saved configuration if any
- **OK** / **Cancel** — start the run, or cancel without running. Enter is OK; Escape is Cancel.

More keys work anywhere in the dialog: **F1** shows Help, **F11** checks the web for a newer version of `urlCheck` (see Updating), **Control+Enter** is OK from any control, **Shift+F1** says the tip for the field you are on, and **F7** lists the dialog's controls so you can move to one. The dialog is built with the Homer Lbc classes, so its text fields also have the Lbc editing keys, which Help lists.

**Default settings** and **Guide** return you to the dialog; Default settings first clears the fields and deletes any saved configuration. If OK finds a problem, such as a missing source or Edge running while Main profile is checked, it says so and returns you to the dialog with your entries kept.

**Note on profiles and privacy.** urlCheck's default — a fresh temporary profile — matches the experience of an anonymous member of the public visiting a site for the first time. The scan captures and analyzes whatever a brand-new visitor would see. Choosing **Main profile** is a deliberate departure from that: pages may be personalized to your account (recommendations tailored to your history, content visible only to you, your name and avatar in the header), and the captured page.htm and screenshot may include that personal information. If you plan to share or publish the output, review it before doing so.

The Browse source and Choose output pickers open at the folder derived from the corresponding text field's current value when that value points to an existing path; otherwise they open at your Documents folder. With **Use configuration** checked, those text fields are pre-populated from your last session, so the pickers naturally pick up where you left off.

If you press OK with an output folder that does not yet exist, urlCheck prompts to create it (default Yes). Choosing No keeps the dialog open with focus on the output field so you can correct it.

When all pages have been processed, a final results dialog summarizes what was done.


### From the command line

Open a Command Prompt in the `urlCheck` program folder, or put that folder on your PATH, and run `urlCheck` with the source as an argument. `urlCheck.cmd` there runs the program in the `exec` folder.

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

- `-a`, `--authenticate` — pause on the first url of each registrable domain so you can authenticate, then press Enter (or click OK) to resume. By default it uses a fresh temporary profile and disconnects Playwright during the pause; combine with `-m` to use your real profile (no disconnect). It turns off `-i`.
- `-f`, `--force` — reuse an existing per-page output folder by emptying it and writing a fresh set of files.
- `-g`, `--gui-mode` — show the parameter dialog.
- `-h`, `--help` — show usage and exit.
- `-i`, `--invisible` — run Microsoft Edge with no visible browser window.
- `-l`, `--log` — also write `urlCheck.log` (UTF-8 with BOM) in the output folder. It adds to the old log unless `-f` is also given, which replaces it.
- `-m`, `--main-profile` — launch Edge with your real (default) Edge profile so saved logins are available. Without `-m`, urlCheck uses a fresh temporary profile so the scan is anonymous. No Edge window may be open.
- `-o <folder>`, `--output-folder <folder>` — write output under `<folder>` (made if missing); the default is the current folder.
- `-u`, `--use-configuration` — read saved settings from `%LOCALAPPDATA%\urlCheck\configs\urlCheck.inix`.
- `-v`, `--version` — show the version and exit.
- `--view-output` — after the run, open the output folder in File Explorer.

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

- **partial** — at least one rule instance fails AND at least one passes for the same criterion.
- **fail** — at least one rule instance fails (no pass on the same criterion).
- **manual** — at least one incomplete result (no fail, no pass).
- **pass** — at least one rule instance passes (no fail, no incomplete).
- **na** — all rule instances are inapplicable to the page.
- **unknown** — no axe rules apply to this criterion. This is the value before a scan.

The Conformance column on the rollup sheet is derived from the per-page Calcs across all included pages (worst-result-wins semantics):

- All `pass` — Supports
- Any `pass` and any `fail` — Partially Supports
- Any `fail` (no pass) — Does Not Support
- Any `manual` (no fail or pass) — Not Evaluated
- Only `na` everywhere — Supports (vacuously satisfied, per WCAG 2.0 Understanding Conformance)
- `unknown` everywhere — Not Evaluated

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

When **Use configuration** is checked in the dialog (or `-u` is on the command line), `urlCheck` reads and writes its settings in:

```
%LOCALAPPDATA%\urlCheck\configs\urlCheck.inix
```

It stores the source field, the output folder, and the option checkboxes, one per line under a `[Settings]` heading, such as `ViewOutput=Yes`. You can edit it in any text editor; a comment you add is kept when `urlCheck` saves. **Default settings** in the dialog deletes the file.

Settings saved by versions before 1.12.0, in `%LOCALAPPDATA%\urlCheck\urlCheck.ini`, are read when there is no `.inix` yet. The old file is removed the first time the new one is saved.

---

## Log files

`urlCheck` always keeps a log of each session, whatever the options:

```
%LOCALAPPDATA%\urlCheck\logs\urlCheck-yyyyMMdd-HHmmss.log
```

Each run gets its own file, named for the moment it started, so sorting the folder by name also sorts it by time. The 30 newest are kept. A log starts with the version, where the program ran from, the Python and Windows versions, the working folder, the command line and every setting, then records each page and any error with its full traceback. If something goes wrong, zip the `logs` folder and send it.

**Log session** (or `-l`) additionally writes `urlCheck.log` in the output folder, beside `ACR.xlsx`, so a log can travel with a set of results. It adds to the log already there, with a blank line between sessions; add **Force replacements** (`-f`) to replace it instead.

Both logs are UTF-8 with a byte-order mark, so Notepad opens them correctly.

---

## Notes

- urlCheck reports the violations the [axe-core](https://github.com/dequelabs/axe-core) engine detects automatically. It does not replace manual testing.
- Local files inside a URL list are loaded as HTML regardless of extension. urlCheck does not validate file contents before loading; if a file is not HTML, Edge may render it unexpectedly. The user is responsible for choosing HTML-renderable files.
- urlCheck waits for the page to finish loading and pauses briefly so late DOM updates are more likely to settle before the scan runs. Pages with very long-running asynchronous content may need a manual retry.

---

## Uninstalling

Use Installed apps in Windows Settings, or the Uninstall shortcut in the `urlCheck` Start-menu group. The uninstaller removes the program, its logs and its saved settings. It does not touch the output folders you made, or any `urlCheck.log` in them.

---

## For developers

How `urlCheck` is built, released and structured is in Developer, in this folder.

## License

MIT License. See `License.htm` at the top of the program folder.
