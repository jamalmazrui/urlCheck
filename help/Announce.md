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

`urlCheck`, `extCheck`, and `2htm` are a small family of free, MIT-licensed Windows tools written by Jamal Mazrui and shared on GitHub. They share control names, dialog layout and command-line spellings wherever the idea is the same, so a person who knows one knows the others.

- [urlCheck](https://github.com/JamalMazrui/urlCheck) — checks web pages with Edge and axe-core, producing per-page reports plus a session-level Accessibility Conformance Report covering all 86 WCAG 2.2 success criteria.
- [extCheck](https://github.com/JamalMazrui/extCheck) — checks the accessibility of `.docx`, `.xlsx`, `.pptx` and `.md` files.
- [2htm](https://github.com/JamalMazrui/2htm) — converts Office documents and other formats to clean, accessible HTML.

Each has a dialog and a command line with the same options, an optional installer with a desktop hotkey (Alt+Control+Shift+U for `urlCheck`, Alt+Control+U for its companion `urlFido`, Alt+Control+2 for `2htm`, Alt+Control+X for `extCheck`), and opt-in recall of the last settings.
