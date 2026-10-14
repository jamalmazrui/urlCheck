---
title: "urlCheck Announcement"
author: "Jamal Mazrui"
---

# urlCheck 1.12.3

`urlCheck` checks web pages for accessibility problems with Microsoft Edge and axe-core, and writes a report for each page plus a draft Accessibility Conformance Report for all of them. Get it from the [urlCheck releases page](https://github.com/JamalMazrui/urlCheck/releases).

## What is new in 1.12

- **A dialog like every Homer dialog.** It is built with the Homer Lbc classes, the same ones the other Homer Tools use, so Control+Enter, Shift+F1, F7 and the Help box work the same everywhere. A Guide button opens the full guide.
- **Update from inside the program.** Press F11 in the dialog to check GitHub for a newer version. If there is one, Enter downloads and starts its installer.
- **A log of every session.** Each run now keeps its own log in `%LOCALAPPDATA%\urlCheck\logs`, so a problem can be explained afterwards even when Log session was off. Log session still also writes `urlCheck.log` beside the results.
- **Settings in a readable file.** Use configuration now saves to `%LOCALAPPDATA%\urlCheck\configs\urlCheck.inix`. Settings from earlier versions are carried over.
- **A tidier installation.** The program is in the `exec` folder and the documents in `help`, as in every Homer Tools program. The installer's results box comes before `urlCheck` starts, so it is never hidden behind it.
- **A full guide and a quick start.** The ReadMe is now a one-page quick start; the full guide is `help\urlCheck.htm`, and Help (F1) opens it.

The full list of changes is in History.

## About the companion accessibility tools

urlCheck, extCheck and 2htm are three free Windows tools for accessibility work. Jamal Mazrui wrote them and shares them on GitHub under the MIT license.

- [urlCheck](https://github.com/JamalMazrui/urlCheck) checks web pages. It opens each page in Microsoft Edge, tests it with axe-core and writes a report for the page. It also writes an Accessibility Conformance Report, as `ACR.xlsx` and `ACR.docx`, that covers all 86 success criteria of WCAG 2.2.
- [extCheck](https://github.com/JamalMazrui/extCheck) checks Word, Excel, PowerPoint and Markdown files. It writes a report of the problems in each file, found by a set of rules that can grow.
- [2htm](https://github.com/JamalMazrui/2htm) turns Office documents and other files into clean, accessible web pages, using Microsoft's own conversion. It can also write plain text, or leave out the images.

The three tools work the same way in the dialog and on the command line. Every option is in both, with the same name and the same default, so a task done by hand can be repeated exactly from a script.

What they share:

- **Accessible in both modes.** Every control in the dialog has its own Alt key, the tab order makes sense, and screen readers announce progress and results. Help is there in both modes.
- **Learn one, know all three.** The tools use the same control names, the same dialog layout and the same command-line options wherever they do the same thing.
- **An installer if you want one.** Each tool has a setup program. It puts the tool in Program Files, adds it to the Start menu, installs its guides and sets a desktop hotkey: Alt+Control+2 for 2htm, Alt+Control+X for extCheck, Alt+Control+Shift+U for urlCheck and Alt+Control+U for urlFido. You can also run a tool without installing it.
- **Many sources at once.** One command can take any number of files, wildcards, folders, web addresses or lists, and handles them one after another.
- **Settings kept, if you choose.** A tool can remember your last choices in the dialog, so a task you do often takes one step.
- **Progress as it happens.** The command line names each file or page as it goes, with a short reason when one fails. The dialog ends with a list of what was done, what failed and what was skipped.
- **Earlier results kept.** A second run keeps earlier results and skips work already done. The force option starts fresh.
- **A session log if you want one,** to look back at what happened.

All three are free under the [MIT license](https://opensource.org/license/mit/). It lets anyone use, change and share the tools, even in business, as long as the copyright notice and the license stay with every copy.
