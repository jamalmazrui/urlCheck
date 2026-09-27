---
title: "urlCheck Developer Notes"
author: "Jamal Mazrui"
---

# urlCheck Developer Notes

How `urlCheck` is built, released and laid out. Since 1.12.0 it is built on the Homer Development Kit (HomerDev), kit 1.43.2 or later, in `C:\HomerDev`.

## The project folder

`C:\urlCheck` mirrors the installed tree:

- At the top: `urlCheck.py` (the whole program), `buildUrlCheck.cmd`, `urlCheck_setup.iss`, `urlCheck.cmd`, `urlCheck.ico`, `requirements.txt`, `accept.inix`, `RepoFiles.txt`, `LocalFiles.txt`, `ReadMe` and `License`.
- `exec` — the built `urlCheck.exe`. Never in git.
- `help` — this document and the others: `urlCheck` (the guide), `Announce`, `Developer`, `History`, `Hotkeys`, each as `.md` and `.htm`.
- `logs` — one log per run of the build or any tool.
- `scripts` — the kit's tools, refreshed from `C:\HomerDev\scripts` by every build.
- `.venv` and `work` — the build's Python environment and PyInstaller's scratch. Never in git.

`version.txt` holds the version and lives only on this machine. The build writes `version.py` from it, and the installer reads it directly.

## The four steps

1. `buildUrlCheck` — steps the version (`buildUrlCheck nobump` keeps it), builds `exec\urlCheck.exe` with PyInstaller, writes each `.htm` from its `.md`, puts the project's files in the Homer encoding, checks that the installer ships every file in `help`, and builds `urlCheck_setup.exe`. Its log is `logs\urlCheck-build-yyyyMMdd-HHmmss.log`.
2. `scripts\push "message"` — rewrites the whitelist `.gitignore` from `RepoFiles.txt`, commits and pushes.
3. `scripts\tidy` and `scripts\tidy --do-it` — the periodic clean.
4. `scripts\release` — runs `scripts\check`, then tags the pushed commit with the version stamped in `urlCheck_setup.exe` and publishes the installer.

Try the fresh build with `urlCheck` (the `urlCheck.cmd` at the top runs `exec\urlCheck.exe`) or `exec\urlCheck.exe -g`.

## What the build fetches

Nothing is fetched by hand. The build installs Python 3.13 (64-bit, machine-wide) with winget when the `py` launcher cannot find it, makes `.venv`, and installs `requirements.txt` and PyInstaller there. It installs pandoc and Inno Setup with winget when they are missing.

Edge is not fetched. `urlCheck` drives the Edge that Windows already has, through Playwright's `channel="msedge"`, and `playwright install msedge` fails on every machine that has Edge.

## The kit's Python modules

`urlCheck.py` imports four of them: `from homer import inix, lbcnet, log, paths`. They are not copied into the project. The build gives PyInstaller `--paths C:\HomerDev` and one `--hidden-import` for each, so `urlCheck.exe` carries the kit's current code.

- `homer.log` keeps the session log in `%LOCALAPPDATA%\urlCheck\logs`. `urlCheck`'s own `logger` class writes every line to it as well as to the optional `urlCheck.log`.
- `homer.paths` finds the per-user folders and, from `exec`, the installed folder with its `help`.
- `homer.inix` reads and writes `configs\urlCheck.inix`.
- `homer.lbcnet` loads `Homer.dll` and gives the dialogs the kit's C# classes.

## The dialog: the kit's C# LbcDialog, from Python

Since 1.12.3 the dialog is not WinForms code of `urlCheck`'s own. It is the kit's C# `LbcDialog`, the class urlFido, bookFido, extCheck and 2htm use. `buildHomerDev` compiles the kit's Elevate, Inix, Lbc, Log, Paths, Say, Util and Web classes into `C:\HomerDev\exec\Homer.dll`; `buildUrlCheck` (with `homerDll=1`) bundles that file into `urlCheck.exe`; and `homer.lbcnet.load()` loads it through pythonnet, sets the thread to a single-threaded apartment, and returns the `Homer` namespace. So the dialog has exactly the focus order, keys, Help box and version check of the C# apps, and a fix to `Lbc.cs` reaches `urlCheck` on its next build.

- `showGuiDialog` builds an `LbcDialog`: two bands for a field and its button, a separator, the seven checkboxes, and `runWithButtons` with OK, Guide, Default settings and Cancel (Lbc adds Help). It loops back to the dialog after Guide, Default settings, or an OK that a check sends back: a missing or invalid source, an output folder the user declined to create, or Main profile with Edge running.
- `lbcnet.strings()` turns a Python list into the .NET `string[]` that `runWithButtons` takes, and `lbcnet.keyHandler()` turns a Python function into the `Func<Keys, bool>` that `commandKey` takes. F11 is claimed that way and answered by the C# `Elevate`, which also supplies the Help box's version section.
- `showFinalGuiMessage` shows long results in an `LbcDialog` holding one read-only multi-line field and OK.
- The Browse and Choose handlers stay Python: pythonnet deadlocks on the modern file and folder pickers, so they use the legacy file dialog and `SHBrowseForFolder` through ctypes.

## Coding style

Camel Type for Python, as `C:\HomerDev\help\CamelType_Python.md` describes. Constants added since the move to the kit carry the `c_` prefix; older ones do not yet.

## Pitfalls already paid for

- PyInstaller resolves `--icon` against the `.spec` file's folder. The build passes the icon's full path, so the `.spec` can live in `work`.
- PyInstaller's bootloader loads a runtime library it cannot unload, so it cannot remove its own `_MEI` folder at exit. `urlCheck` removes the leftovers of earlier runs at startup, and `urlCheck.cmd` sends the bootloader's one warning to `nul`.
- pythonnet deadlocks on the modern file and folder pickers, so `urlCheck` uses the legacy file dialog and `SHBrowseForFolder` through ctypes.
- A copy of `exec\urlCheck.exe` that is running cannot be replaced. The build says so and stops; it never closes a program. An installed copy may stay open.
