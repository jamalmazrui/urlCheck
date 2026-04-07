# urlCheck

**Author:** Jamal Mazrui  
**License:** MIT

`urlCheck` is a Windows command-line tool that checks web pages for accessibility problems. It opens a page in a browser, runs the axe-core testing engine, and saves a set of output files in a new folder named after the page title.

The whole urlCheck project may be downloaded in a single zip archive from:

<http://GitHub.com/JamalMazrui/urlCheck/archive/main.zip>

---

## What you need

- Windows 10 or later
- An internet connection during each scan

---

## How to use it

Use `urlCheck.cmd` to run the program. The `.cmd` wrapper suppresses a harmless internal message that can appear when the program closes.

**Show help:**

```cmd
urlCheck.cmd -h
```

**Show version:**

```cmd
urlCheck.cmd -v
```

**Check a web page:**

```cmd
urlCheck.cmd https://example.com
```

**Check a page using just the domain name** (`https://` is added automatically):

```cmd
urlCheck.cmd microsoft.com
```

**Check a local HTML file:**

```cmd
urlCheck.cmd "C:\work\sample.html"
```

**Check a list of URLs from a text file:**

```cmd
urlCheck.cmd urls.txt
```

The program prints its name and version right away, then shows progress as it works. After each scan it prints a short summary of what it found.

---

## Scanning a list of URLs

If you give `urlCheck` a text file, it treats each line as a URL to check. Blank lines and lines that start with `#` are skipped.

Each URL gets its own output folder. The program does not open `report.htm` automatically when running a list.

**Example `urls.txt`:**

```
# Pages to check - April 2026
https://example.com
https://example.org/about
https://example.net/contact
```

At the end, the program tells you how many pages were checked and how many had errors.

---

## Output files

Each scan creates a folder named after the page title. Inside that folder you will find these files:

| File | What it contains |
|---|---|
| `report.htm` | A readable accessibility report with headings and links |
| `report.csv` | A spreadsheet-ready list of violations, one row per issue |
| `report.xlsx` | An Excel workbook with multiple sheets of results |
| `results.json` | The full raw data from the accessibility scan, including all metadata |
| `page.yaml` | The ARIA accessibility tree of the page |
| `page.htm` | A saved copy of the page source with styles included |
| `page.png` | A full-page screenshot |

If an error occurs, the program writes an `error.txt` file to the folder instead.

### report.htm

This is the main report. It has a table of contents with anchor links so you can jump to any section. Sections include scan details, a count of issues by type, the most common problems, and a full list of violations with HTML snippets and fix guidance. At the end there is a glossary and links to public accessibility resources.

### report.csv

Each row is one failing element on the page. Columns include the rule name, impact level, a description of the problem, the CSS selector of the failing element, and the HTML of that element. All fields are fully quoted so the file opens cleanly in Excel, even when HTML attributes contain quotes.

Columns included:

`scanTimestampUtc`, `pageTitle`, `pageUrl`, `browserVersion`, `axeSource`, `outcome`, `ruleId`, `impact`, `description`, `help`, `helpUrl`, `tags`, `wcagRefs`, `standardsRefs`, `ruleNodeCount`, `ruleNodeIndex`, `target`, `html`, `failureSummary`

### report.xlsx

An Excel workbook with four sheets:

- **Metadata** — URL, page title, browser version, scan time, and file names
- **Summary** — counts of violations by impact level and most common rules
- **Results** — same data as `report.csv`, one row per issue
- **Glossary** — definitions of terms and a list of steps the program follows

### results.json

The complete raw output from the accessibility scan, plus metadata. This file includes all four outcome types: violations, incomplete results that need manual review, passes, and inapplicable rules.

### page.yaml

The ARIA accessibility tree of the page. The output is a YAML file representing the hierarchical structure of accessible elements, showing roles, accessible names, ARIA attributes such as level, checked, and expanded, and text content. Only nodes visible to assistive technologies are included. The file is encoded as UTF-8 with a byte-order mark (BOM).

### page.htm

A snapshot of the page as the browser saw it. External CSS and JavaScript files are inlined where the program can retrieve them.

### page.png

A full-page screenshot taken at a 1600 x 1440 viewport.

---

## Notes

- `urlCheck` only reports violations automatically. It does not replace manual testing.
- If the server that hosts the accessibility testing engine is not reachable, the scan will fail.
- The tool works with web URLs and local `.html`, `.htm`, and `.xhtml` files.

---

## Development

This section is for developers who want to build `urlCheck.exe` from source.

### Prerequisites

- Python 3.14 or later from [python.org](https://www.python.org/downloads/)
- The source files from this repository

### Install build dependencies

```cmd
py -3.14 -m pip install --upgrade pip
py -3.14 -m pip install --upgrade -r requirements.txt
py -3.14 -m pip install "pyinstaller>=6.19.0"
```

### Set up the browser driver

```cmd
py -3.14 -m playwright install msedge
```

### Build the executable

```cmd
py -3.14 -m PyInstaller --clean --noconfirm --onefile --name urlCheck urlCheck.py
```

The result is at `dist\urlCheck.exe`.

### Automated build

Run the included script, which installs dependencies, builds the executable, and copies `urlCheck.cmd` into the `dist\` folder:

```cmd
buildUrlCheck.cmd
```

Distribute both files together:

```
dist\urlCheck.exe
dist\urlCheck.cmd
```

### Running from source

To run without building the executable:

```cmd
py -3.14 urlCheck.py microsoft.com
```

---

## License

MIT License. See `license.htm`.
