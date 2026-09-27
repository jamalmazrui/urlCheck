---
title: "urlCheck ReadMe"
author: "Jamal Mazrui"
---

# urlCheck ReadMe

`urlCheck` checks web pages for accessibility problems. It opens each page in Microsoft Edge, runs the [axe-core](https://github.com/dequelabs/axe-core) testing engine, and saves a report in a new folder named after the page title. At the end of each run it also writes a draft Accessibility Conformance Report for all the pages checked.

This is the quick start. The full guide is `help\urlCheck.htm`, which Help (F1) in the program also opens.

## Install

1. Download `urlCheck_setup.exe` from the [urlCheck releases page](https://github.com/JamalMazrui/urlCheck/releases).
2. Run it. It asks for administrator rights, because it installs for everyone on the computer.
3. On the last page, leave **Launch urlCheck now** checked and press Finish. Read the results box, then close it; `urlCheck` opens.

You need Windows 10 or 11, 64-bit. Edge is already part of Windows, and nothing else needs installing.

## Check a page

1. Press **Alt+Control+Shift+U** from anywhere in Windows. The `urlCheck` dialog opens with focus in **Source urls**.
2. Type a web address, such as `https://example.com`, or several separated by spaces.
3. Press Enter.

Edge opens, the page is checked, and a results box says what was done. The report is in a folder named after the page, inside the output folder (your Documents folder unless you choose another).

## From the command line

In a Command Prompt in the program folder:

```cmd
urlCheck https://example.com
urlCheck urls.txt -o reports --view-output
urlCheck --help
```

## Keys in the dialog

- **Alt** with an underlined letter moves to that control.
- **Enter** starts the check; **Escape** closes the dialog.
- **F1** shows Help and offers the full guide.
- **F11** checks the web for a newer version of `urlCheck` and offers to install it.

All the keys are listed in `help\Hotkeys.htm`.

## When something goes wrong

Every run keeps a log in `%LOCALAPPDATA%\urlCheck\logs`, one file per run. Zip that folder and send it with a description of what happened.

## License

`urlCheck` is free and open source under the MIT License. See `License.htm`.
