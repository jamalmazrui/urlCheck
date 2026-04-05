# Batch Checks of Accessibility for Public Web Pages and Popular File Formats

I am pleased to share two free, open-source accessibility testing tools I have just released on GitHub. Each is a single, self-contained executable that runs from any folder on a Windows computer — no installation required, and no dependencies to manage.

## extCheck — Check Office and Markdown Files

https://github.com/jamalmazrui/extCheck

extCheck checks Microsoft Word, Excel, PowerPoint, and Pandoc Markdown files against 73 accessibility rules drawn from the Microsoft Office Accessibility Checker and the axe-core WCAG 2.2 rule set. Results appear in the console and in a CSV report, with each issue described along with remediation guidance.

## urlCheck — Check Web Pages

https://github.com/jamalmazrui/urlCheck

urlCheck opens a web page in Microsoft Edge, runs the axe-core engine, and saves an HTML report, a CSV file, an Excel workbook, and a JSON results file — all in a named folder. It works with public URLs, local HTML files, and batch lists of URLs.

---

Because each tool is a portable single file, it integrates naturally into developer pipelines: call it synchronously, then parse the machine-readable CSV or JSON output as part of a larger automated workflow.

Both tools are designed for users comfortable with the Windows command line. If you work in accessibility testing, document production, or web quality assurance, I hope you find them useful. Feedback and contributions are welcome.

\#accessibility \#wcag \#a11y \#assistivetechnology \#commandlinetools \#opensource