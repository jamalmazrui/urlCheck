# Tutorials

## Contents

- [1 - User Interface](#1-user-interface)
- [2 - Install and Launch](#2-install-and-launch)
- [3 - Check One Web Page](#3-check-one-web-page)
- [4 - Read the Workbook and Screenshots](#4-read-the-workbook-and-screenshots)
- [5 - Check a List of Pages](#5-check-a-list-of-pages)
- [6 - Check Pages Behind a Sign In](#6-check-pages-behind-a-sign-in)
- [7 - Draft a Conformance Report](#7-draft-a-conformance-report)
- [8 - Check from the Command Line](#8-check-from-the-command-line)
- [9 - Conclusion](#9-conclusion)
- [1 - User Interface](#1-user-interface)
- [2 - Install and Launch](#2-install-and-launch)
- [3 - Check One Web Page](#3-check-one-web-page)
- [4 - Read the Workbook and Screenshots](#4-read-the-workbook-and-screenshots)
- [5 - Check a List of Pages](#5-check-a-list-of-pages)
- [6 - Check Pages Behind a Sign In](#6-check-pages-behind-a-sign-in)
- [7 - Draft a Conformance Report](#7-draft-a-conformance-report)
- [8 - Check from the Command Line](#8-check-from-the-command-line)
- [9 - Conclusion](#9-conclusion)
- [0 - Overview](#0-overview)
- [1 - User Interface](#1-user-interface)
- [2 - Install and Launch](#2-install-and-launch)
- [3 - Check One Web Page](#3-check-one-web-page)
- [4 - Read the Workbook and Screenshots](#4-read-the-workbook-and-screenshots)
- [5 - Check a List of Pages](#5-check-a-list-of-pages)
- [6 - Check Pages Behind a Sign In](#6-check-pages-behind-a-sign-in)
- [7 - Draft a Conformance Report](#7-draft-a-conformance-report)
- [8 - Check from the Command Line](#8-check-from-the-command-line)
- [9 - Conclusion](#9-conclusion)
- [1. Where to Go Next](#1-where-to-go-next)

<!-- walkthrough: written by makeTutorials.py, do not edit between the markers -->

## 0 - Overview

What urlCheck is, what it writes, and the ten walks that teach it.

**Before you start:** Nothing to set up; this walk only listens.

### Step 1

Welcome. This is the first of ten short walks through urlCheck. I am the host; the other voice is the screen reader, speaking as it would on your own computer.

### Step 2: Insert+Up Arrow

Two reader keys help in every walk. If a line goes by too fast, Insert plus Up Arrow says it again.

Screen reader:

- Source urls edit

Confirm the wording with a live run: buildTutorials -live.

### Step 3: Insert+Tab

And whenever you are unsure where you are, Insert plus Tab says what has the focus.

Screen reader:

- Source urls edit, blank

Confirm the wording with a live run: buildTutorials -live.

### Step 4

urlCheck checks web pages for accessibility problems. It opens each page in Microsoft Edge, as a person would, and runs axe-core, the open-source checking engine many testers use.

### Step 5

For each page it writes a report you can read in a browser, a workbook you can sort, screenshots of each problem, and the page's own accessibility tree.

### Step 6

At the end of every run it drafts an Accessibility Conformance Report, one row for each criterion of WCAG 2.2, ready to review and complete.

### Step 7

It runs from a small dialog opened from the desktop, or from the command line, where one line can check a whole list of pages.

### Step 8

Here are the walks. One, the dialog. Two, installing. Three, checking one page. Four, the workbook and screenshots. Five, a list of pages.

### Step 9

Six, pages behind a sign in. Seven, the conformance report. Eight, the command line. Nine, the conclusion, a glossary, and where to get help.

### Step 10

An automated check finds many problems but not all, so the conformance report also lists the manual tests each criterion needs.

### Step 11

Each walk is a few minutes long and builds on the last. Walk one comes next.

**Something to try:** Choose one web page you care about. Walk three checks it.

## 1 - User Interface

The urlCheck dialog: its fields and boxes, their underlined letters, and the keys that work anywhere in it, each shown as the reader speaks it.

**Before you start:** urlCheck installed; the dialog open with Alt+Control+Shift+U.

### Step 1

One want: to move around the urlCheck dialog quickly, knowing what every box does before the first check.

### Step 2: Alt+Control+Shift+U

I press the desktop key. The dialog opens on its first field.

Screen reader:

- urlCheck dialog
- Source urls edit

Confirm the wording with a live run: buildTutorials -live.

### Step 3

Source urls names what to check: one address, several separated by spaces, or the path of a text file listing them. Its letter is S.

### Step 4: Tab

Tab moves to the next control, a button that picks such a file.

Screen reader:

- Browse source button

Confirm the wording with a live run: buildTutorials -live.

### Step 5: Alt+O

Alt with an underlined letter jumps straight to a control. Output folder is O.

Screen reader:

- Output folder edit

Confirm the wording with a live run: buildTutorials -live.

### Step 6

Reports go into that folder, one subfolder for each page; blank means the current folder. Choose output, C, picks one instead of typing.

### Step 7: Alt+A

Then the boxes. Authenticate credentials, A, pauses at the first page of each site so you can sign in.

Screen reader:

- Authenticate credentials check box, not checked

Confirm the wording with a live run: buildTutorials -live.

### Step 8: Alt+M

Main profile, M, uses your everyday Edge profile, with its saved sign ins.

Screen reader:

- Main profile check box, not checked

Confirm the wording with a live run: buildTutorials -live.

### Step 9: Alt+I

Invisible mode, I, runs Edge with no window at all.

Screen reader:

- Invisible mode check box, not checked

Confirm the wording with a live run: buildTutorials -live.

### Step 10: Alt+F

Force replacements, F, checks a page again even when its folder already holds a report.

Screen reader:

- Force replacements check box, not checked

Confirm the wording with a live run: buildTutorials -live.

### Step 11: Alt+V

View output, V, opens the folder at the end; Log session, L, writes a log; Use configuration, U, remembers these choices.

Screen reader:

- View output check box, not checked

Confirm the wording with a live run: buildTutorials -live.

### Step 12: Space

Space checks a box, and Space again clears it.

Screen reader:

- checked

Confirm the wording with a live run: buildTutorials -live.

### Step 13: Space

Space.

Screen reader:

- not checked

Confirm the wording with a live run: buildTutorials -live.

### Step 14: Shift+F1

Shift plus F1 says the tip for the control you are on.

Screen reader:

- View output: open the output folder when the run is done

Confirm the wording with a live run: buildTutorials -live.

### Step 15: F1

F1 shows Help, with every control and every key.

Screen reader:

- urlCheck Help dialog

Confirm the wording with a live run: buildTutorials -live.

### Step 16: Escape

Escape closes Help and returns to the same place.

Screen reader:

- View output check box, not checked

Confirm the wording with a live run: buildTutorials -live.

### Step 17: F7

F7 lists every control; choose one and press Enter to move there.

Screen reader:

- Controls list

Confirm the wording with a live run: buildTutorials -live.

### Step 18: Enter

I choose Source urls.

Screen reader:

- Source urls edit

Confirm the wording with a live run: buildTutorials -live.

### Step 19: Tab

Enter starts the check from any field, and Control plus Enter from any control. Escape closes the dialog without checking.

Screen reader:

- Browse source button

Confirm the wording with a live run: buildTutorials -live.

### Step 20: Alt+G

Guide, G, opens the full guide and returns here; Default settings, D, clears every field and box.

Screen reader:

- Guide button

Confirm the wording with a live run: buildTutorials -live.

### Step 21

The text fields remember earlier answers: Up Arrow in Source urls brings back the last addresses checked.

### Step 22

Each box's state is saved the moment you answer it when Use configuration is ticked, so closing the dialog by Escape loses nothing.

### Step 23: Alt+D

Default settings clears every field and box and deletes the saved configuration, so it is for starting fresh, not for tidying.

Screen reader:

- Default settings button

Confirm the wording with a live run: buildTutorials -live.

### Step 24

That is the dialog: what to check, where reports go, seven boxes, and a key for each. The next walk installs urlCheck.

**Something to try:** Open the dialog, press F7, choose a control, then press Shift+F1 to hear its tip.

## 2 - Install and Launch

Installing urlCheck for everyone on the computer, opening it from anywhere, and keeping it current.

**Before you start:** urlCheck_setup.exe downloaded from the urlCheck releases page on GitHub.

### Step 1

One want: urlCheck on this computer, opened with one key, ready to check pages.

### Step 2: Enter

I open the downloaded setup program. It installs for everyone, so Windows asks for administrator rights; I say yes.

Screen reader:

- Setup - urlCheck dialog
- Welcome to the urlCheck Setup Wizard

Confirm the wording with a live run: buildTutorials -live.

### Step 3: Enter

The welcome page summarizes the license: urlCheck is free and open source, under the MIT license. Enter takes Next.

Screen reader:

- Select Destination Location

Confirm the wording with a live run: buildTutorials -live.

### Step 4: Enter

The default folder in Program Files is right; an update goes there without asking.

Screen reader:

- Ready to Install

Confirm the wording with a live run: buildTutorials -live.

### Step 5: Enter

Enter installs it.

Screen reader:

- Installing

Confirm the wording with a live run: buildTutorials -live.

### Step 6

The last page offers to launch urlCheck now.

Screen reader:

- Launch urlCheck now check box, checked

Confirm the wording with a live run: buildTutorials -live.

### Step 7: Enter

I press Finish. A short box says what was installed and where the logs are.

Screen reader:

- urlCheck setup complete

Confirm the wording with a live run: buildTutorials -live.

### Step 8: Enter

When I close it, urlCheck opens on its dialog.

Screen reader:

- urlCheck dialog
- Source urls edit

Confirm the wording with a live run: buildTutorials -live.

### Step 9

From now on, Alt plus Control plus Shift plus U opens it from anywhere in Windows, and the Start menu has it too.

### Step 10

urlCheck drives Microsoft Edge, which Windows 10 and 11 already have, so nothing else needs installing to check a page.

### Step 11

It keeps a browser profile of its own for checking, separate from your everyday Edge, unless you choose Main profile.

### Step 12

Your settings and logs live in your own application data, so an update never touches them.

### Step 13: F11

F11 in the dialog asks GitHub for a newer version; when there is one, Enter installs it.

Screen reader:

- Checking for a newer version

Confirm the wording with a live run: buildTutorials -live.

### Step 14: Enter

If this copy is the newest, No is the default, and Enter returns to the dialog.

Screen reader:

- urlCheck dialog

Confirm the wording with a live run: buildTutorials -live.

### Step 15

Uninstalling, from Installed apps in Windows Settings, leaves every report you made where you saved it.

### Step 16

If setup ever stops partway, its log says why, and running it again finishes the job.

### Step 17

The full guide is in the help folder of the installation, readable before the first check.

### Step 18

Each run writes a session log in your own application data, whatever the dialog says, so a problem can be traced after the fact.

### Step 19

urlCheck's own browser profile is wiped of cookies between runs unless you choose Main profile, so one site's sign in never leaks into another check.

### Step 20

A run that cannot find Edge says so in plain words, with what to install, rather than failing silently.

### Step 21

The program, its documents and its scripts all sit in Program Files, apart from anything it writes for you.

### Step 22

Installed, opened with one key, and kept current. The next walk checks a page.

**Something to try:** After installing, open the dialog with its desktop key and press F11.

## 3 - Check One Web Page

Checking a single web page, and reading the report urlCheck writes for it.

**Before you start:** urlCheck installed, and a web page you want to check.

### Step 1

One want: to know whether a site's home page has accessibility problems, before writing to its owner about them.

### Step 2: Alt+Control+Shift+U

I open the dialog. The focus is on Source urls.

Screen reader:

- urlCheck dialog
- Source urls edit

Confirm the wording with a live run: buildTutorials -live.

### Step 3

I type the address. The https part is optional; urlCheck adds it.

Screen reader:

- example.org

Confirm the wording with a live run: buildTutorials -live.

### Step 4: Alt+V

I leave Output folder blank, so reports go into the current folder, and I tick View output.

Screen reader:

- View output check box, checked

Confirm the wording with a live run: buildTutorials -live.

### Step 5: Enter

Enter starts the check. Edge opens the page, waits for it to settle, then runs axe-core.

Screen reader:

- Checking example.org

Confirm the wording with a live run: buildTutorials -live.

### Step 6

A short box reports the totals when it is done.

Screen reader:

- 1 page checked
- OK button

Confirm the wording with a live run: buildTutorials -live.

### Step 7: Enter

Enter closes it, and File Explorer opens on the output folder. There is a subfolder named after the page's title.

Screen reader:

- Example Domain

Confirm the wording with a live run: buildTutorials -live.

### Step 8: Enter

Enter goes into it. Inside are the report, the workbook, the screenshots and the page's own files.

Screen reader:

- report dot h t m

Confirm the wording with a live run: buildTutorials -live.

### Step 9: Enter

Enter on report dot h t m opens the report in the browser.

Screen reader:

- urlCheck report, heading level 1

Confirm the wording with a live run: buildTutorials -live.

### Step 10: Down Arrow

It begins with a summary: the page, when it was checked, and how many rules were violated, passed, or could not be decided.

Screen reader:

- Violations: 3. Passes: 24. Incomplete: 2.

Confirm the wording with a live run: buildTutorials -live.

### Step 11: H

H moves to the violations, each rule a heading of its own.

Screen reader:

- Images must have alternate text, heading level 2

Confirm the wording with a live run: buildTutorials -live.

### Step 12: Down Arrow

Under each rule: what it means, how serious it is, which WCAG criterion it maps to, and how many elements break it.

Screen reader:

- Impact: critical. WCAG 1.1.1

Confirm the wording with a live run: buildTutorials -live.

### Step 13: Down Arrow

Then each element that breaks it, by its place in the page, with the advice axe gives for fixing it.

Screen reader:

- Fix: give the image an alt attribute

Confirm the wording with a live run: buildTutorials -live.

### Step 14

Incomplete means axe could not decide by itself, such as colour contrast over a picture; those are for a person to check.

### Step 15

Passes are listed too, so the report shows what the page does well as well as what it does wrong.

### Step 16: Alt+F

Checking the same page again keeps the first report, since its folder exists; Force replacements checks it afresh.

Screen reader:

- Force replacements check box, checked

Confirm the wording with a live run: buildTutorials -live.

### Step 17

A violation count is a start, not a verdict: walk seven turns these results into a conformance report with the manual tests still to do.

### Step 18

Impact comes in four levels, critical, serious, moderate and minor, which is the order to fix things in.

### Step 19: Insert+F6

Insert plus F6 lists the report's headings, one for each violated rule, a quick table of contents to the problems.

Screen reader:

- Headings list

Confirm the wording with a live run: buildTutorials -live.

### Step 20: Down Arrow

Each element is described by its place in the page and a snippet of its code, which is what a developer needs to find it.

Screen reader:

- img src logo dot png

Confirm the wording with a live run: buildTutorials -live.

### Step 21: K

The WCAG number beside each rule is a link to that criterion's explanation, for anyone who wants the why behind the rule.

Screen reader:

- link, WCAG 1.1.1

Confirm the wording with a live run: buildTutorials -live.

### Step 22

A report from before and after a fix, side by side, is the simplest proof that the fix worked.

### Step 23

One page, one report, ready to read. The next walk opens the workbook and its screenshots.

**Something to try:** Check the home page of a site you use often, and count its violations.

## 4 - Read the Workbook and Screenshots

The workbook urlCheck writes for each page: its sheets, its screenshot links, and the other files in the page's folder.

**Before you start:** A page checked in walk three, and Excel installed.

### Step 1

One want: the same results in a form I can sort and share with a developer, with a picture of each problem for sighted colleagues.

### Step 2: Enter

In the page's folder, report dot x l s x is the workbook. I open it in Excel.

Screen reader:

- report dot x l s x - Excel

Confirm the wording with a live run: buildTutorials -live.

### Step 3

It has a sheet for each kind of result: violations, passes, incomplete, and rules that did not apply.

### Step 4: Control+Page Down

Control plus Page Down moves from sheet to sheet, and the screen reader says each sheet's name.

Screen reader:

- Passes

Confirm the wording with a live run: buildTutorials -live.

### Step 5: Control+Page Up

Control plus Page Up returns to the violations.

Screen reader:

- Violations

Confirm the wording with a live run: buildTutorials -live.

### Step 6: Down Arrow

Each row is one rule and one element, with columns for the rule, its impact, the WCAG criterion, the element and the fix.

Screen reader:

- image-alt, critical, 1.1.1

Confirm the wording with a live run: buildTutorials -live.

### Step 7

Sorting by impact, most critical first, puts the worst problems at the top for a developer's list.

### Step 8

The Results sheet has an Image column: each cell is a link to a screenshot of that one element.

Screen reader:

- Image, violations backslash image-001 dot p n g, link

Confirm the wording with a live run: buildTutorials -live.

### Step 9

Control plus K, or Enter on the link, opens the picture, which a sighted colleague can see at once.

### Step 10

The screenshots are in the page's violations folder, numbered in order, image 001, image 002 and so on.

### Step 11

The links are relative, so the whole page folder can be zipped, emailed or moved, and they still work wherever it lands.

### Step 12

When axe's description of an element cannot be found on the page, such as an element hidden or inside a shadow tree, no screenshot is taken, and the row says nothing is attached.

### Step 13

Besides the workbook, the folder holds page dot p n g, a picture of the whole page, and page dot h t m, its source as it was checked.

### Step 14

Results dot j s o n holds everything in a form other programs read, for anyone building on urlCheck's results.

### Step 15

And page dot y a m l is the page's accessibility tree, roles and names as a screen reader receives them, readable in any text editor.

### Step 16

Reading that tree is a quick way to hear how a page is built, without the page's visual layout in the way.

### Step 17: Alt+Down Arrow

Excel's AutoFilter narrows the sheet to one impact or one rule, which turns a long list into a short one.

Screen reader:

- Filter, Impact

Confirm the wording with a live run: buildTutorials -live.

### Step 18

Each sheet has its header row in the first row, so the screen reader announces the column name with every cell.

### Step 19

The Passes sheet is worth a look: it shows the rules a page already meets, which is good news to share alongside the problems.

### Step 20

A developer can be sent the workbook alone; its screenshot links need the violations folder beside it, so zip the page's whole folder.

### Step 21

A report to read, a workbook to sort, a picture of each problem: one folder serves blind and sighted reviewers alike. The next walk checks a list of pages.

**Something to try:** Open one violation's screenshot from the workbook and share it with a sighted colleague.

## 5 - Check a List of Pages

Checking many pages in one run, from addresses typed together or listed in a text file.

**Before you start:** A text file listing several addresses, one to a line.

### Step 1

One want: to check every important page of a site, not just its home page, in one run I can leave to itself.

### Step 2

Several addresses can go into Source urls together, separated by spaces.

Screen reader:

- example.org example.org slash about

Confirm the wording with a live run: buildTutorials -live.

### Step 3

For more than a few, a text file is easier: one address to a line, saved anywhere, with any name.

### Step 4

Local web pages can be listed too, by their paths, for checking a site before it is published.

### Step 5: Alt+B

Browse source, B, picks that file.

Screen reader:

- Open dialog
- File name edit

Confirm the wording with a live run: buildTutorials -live.

### Step 6: Enter

I choose my list and press Enter. urlCheck recognizes a list by what is in it, not by its extension.

Screen reader:

- Source urls edit
- C colon backslash Lists backslash Site pages dot t x t

Confirm the wording with a live run: buildTutorials -live.

### Step 7: Alt+O

I put the reports in a folder of their own, Site Check.

Screen reader:

- Output folder edit

Confirm the wording with a live run: buildTutorials -live.

### Step 8: Alt+L

Log session, L, keeps a record of the whole run.

Screen reader:

- Log session check box, checked

Confirm the wording with a live run: buildTutorials -live.

### Step 9: Enter

Enter, and the run begins, one page at a time.

Screen reader:

- Checking 2 of 12: About us

Confirm the wording with a live run: buildTutorials -live.

### Step 10

Each page gets its own subfolder, named after its title, with its own report and workbook.

### Step 11

The result box gives the totals for the run.

Screen reader:

- 12 pages checked

Confirm the wording with a live run: buildTutorials -live.

### Step 12

A page that would not load is named in the log with the reason, and the others are still checked.

### Step 13

Running the list again tomorrow skips pages whose folders exist, so only new pages are checked; Force replacements checks them all again.

### Step 14: Alt+U

Use configuration, U, remembers the list and the folder, so the next run is one key.

Screen reader:

- Use configuration check box, checked

Confirm the wording with a live run: buildTutorials -live.

### Step 15

Comparing the pages' reports shows which problems belong to the whole site, such as a menu, and which to one page.

### Step 16

A problem found on every page is usually one fix in a shared template, which is worth telling a developer first.

### Step 17

Each page's folder is named after its title, adjusted for the file system, so two pages with the same title get folders told apart.

### Step 18

A list can mix addresses of several sites; each is checked in turn, and sign ins, when asked for, are per site.

### Step 19

Blank lines and lines starting with a number sign are skipped, so a list can carry notes about what each address is for.

### Step 20

A long list can run while you do other work; Edge opens and closes on its own window, which is best left alone.

### Step 21

The conformance report at the end covers every page of the run, which is where walk seven picks up.

### Step 22

A whole site, checked in one run, one folder per page. The next walk checks pages behind a sign in.

**Something to try:** Check five pages of one site from a list, then compare their violation counts.

## 6 - Check Pages Behind a Sign In

Checking pages that need you to sign in: pausing to authenticate, or using your everyday Edge profile and its saved sign ins.

**Before you start:** An account on a site whose inner pages you want to check.

### Step 1

One want: to check the pages people use after signing in, such as an account page, not only the public ones.

### Step 2

A page behind a sign in shows a sign in form to a browser that has not signed in, so checking it plainly checks the wrong page.

### Step 3: Alt+A

Authenticate credentials, A, solves this: urlCheck pauses at the first page of each site, so you can sign in yourself.

Screen reader:

- Authenticate credentials check box, not checked

Confirm the wording with a live run: buildTutorials -live.

### Step 4: Space

Space checks it.

Screen reader:

- checked

Confirm the wording with a live run: buildTutorials -live.

### Step 5: Enter

I put the account page's address in Source urls and press Enter. Edge opens, and urlCheck waits.

Screen reader:

- Sign in, then press OK to continue

Confirm the wording with a live run: buildTutorials -live.

### Step 6: Enter

I sign in in the Edge window as usual, then return to urlCheck's box and press Enter.

Screen reader:

- Checking account page

Confirm the wording with a live run: buildTutorials -live.

### Step 7

From then on, the rest of that site's pages in the run are checked as a signed in user.

### Step 8

urlCheck pauses once for each site, so a list covering two sites asks twice.

### Step 9: Alt+M

The other way is Main profile, M: urlCheck uses your everyday Edge profile, with the sign ins it already holds.

Screen reader:

- Main profile check box, not checked

Confirm the wording with a live run: buildTutorials -live.

### Step 10

That needs no pause, but Edge must be closed first, since one profile cannot be open twice.

### Step 11

Without either, urlCheck uses a fresh profile of its own each time, which is the right choice for public pages: no cookies, no surprises.

### Step 12

Your password never passes through urlCheck; you type it into Edge, as on any other day.

### Step 13

Leave Invisible mode off for a run that asks you to sign in, since the sign in needs a window to type into.

### Step 14

The reports for these pages are written like any others, so an account page's problems sit beside the public pages' in the same folder.

### Step 15

Pages behind a sign in are often the ones people use most, so checking them matters most.

### Step 16

When the run is done, sign out in Edge if the account matters; urlCheck's own profile keeps the session otherwise.

### Step 17

A site that signs you out after a few minutes may need its list split into shorter runs, each with its own pause.

### Step 18

Two factor codes work as usual, since you are the one signing in, at your own pace, before pressing OK.

### Step 19

Checking the sign in page itself needs no account, and is often where accessibility problems stop people first.

### Step 20

Main profile suits a quick check of a page you already use daily; Authenticate suits a planned review with a clean profile.

### Step 21

Signed in once, checked like the rest. The next walk drafts a conformance report.

**Something to try:** Check one page behind a sign in, using Authenticate credentials.

## 7 - Draft a Conformance Report

The ACR workbook and document urlCheck writes at the end of every run: one row per WCAG 2.2 criterion, with a conformance verdict and the manual tests still to do.

**Before you start:** A run of one or more pages, from walk three or five.

### Step 1

One want: a draft Accessibility Conformance Report for a product, the document buyers ask for, without starting from a blank template.

### Step 2

At the end of every run, urlCheck writes ACR dot x l s x and ACR dot d o c x in the output folder, beside the page folders.

Screen reader:

- ACR dot d o c x

Confirm the wording with a live run: buildTutorials -live.

### Step 3

Together they draft a report for every page checked, mapped to WCAG 2.2.

### Step 4: Enter

I open the workbook. Its first sheet, Conformance Report, has one row for each criterion of WCAG 2.2, A, AA and AAA.

Screen reader:

- Conformance Report

Confirm the wording with a live run: buildTutorials -live.

### Step 5: Down Arrow

The first column names the criterion and its level, linked to the W3C's quick reference for it.

Screen reader:

- 1.1.1 Non-text Content (A)

Confirm the wording with a live run: buildTutorials -live.

### Step 6: Tab

The next is a one-sentence summary of what the criterion asks.

Screen reader:

- Summary: give text alternatives for non-text content

Confirm the wording with a live run: buildTutorials -live.

### Step 7: Tab

Then Conformance, the verdict in VPAT terms: Supports, Partially Supports, Does Not Support, or Not Evaluated.

Screen reader:

- Does Not Support

Confirm the wording with a live run: buildTutorials -live.

### Step 8

Below the verdict, the cell lists the pages whose results led to it, so a reviewer can trace it.

### Step 9: Tab

Then Manual: numbered steps to test what no automated check can, such as whether alternative text is actually right.

Screen reader:

- Manual: 1. Check each image's text describes its purpose.

Confirm the wording with a live run: buildTutorials -live.

### Step 10

Not Evaluated means no automated rule covers that criterion; the manual steps are the only test, and a person must do them.

### Step 11

Remarks are left for you, the reviewer, to explain each verdict in your own words.

### Step 12

The document, ACR dot d o c x, presents the same rows in the form buyers expect, ready to edit in Word.

### Step 13

Sheets after the first hold each page's own rows, so a verdict can be followed back to the page that caused it.

### Step 14

A draft is a start: an honest report needs the manual tests done and the remarks written before it is shared.

### Step 15

Run again after fixes, and the draft is rewritten from the new results, so the report keeps pace with the work.

### Step 16

Level A criteria are the baseline, AA is what most laws and buyers ask for, and AAA goes further; the workbook lists all three so a reviewer chooses the scope.

### Step 17

The verdicts come from axe's results mapped to each criterion, so a Supports verdict means no violation was found, not that every manual test passed.

### Step 18: Alt+Down Arrow

Filtering the Conformance column to Does Not Support gives the list of work to do first.

Screen reader:

- Filter, Conformance

Confirm the wording with a live run: buildTutorials -live.

### Step 19

Each criterion's link to the quick reference opens the W3C's explanation and techniques, for writing an accurate remark.

### Step 20

Keep each draft beside the page folders it came from, so a buyer's question about a verdict can always be traced to the evidence.

### Step 21

Every criterion, a verdict, and the tests still to do. The last task walk runs it all from the command line.

**Something to try:** Open ACR dot docx, and find one criterion marked Not Evaluated.

## 8 - Check from the Command Line

Checking pages from a command prompt or a batch file, with the same options as the dialog, invisible mode included.

**Before you start:** A command prompt.

### Step 1

One want: to check a site's pages every week, without opening the dialog, and to know when something new breaks.

### Step 2

At a command prompt, urlCheck followed by an address checks it.

Screen reader:

- urlCheck example.org

Confirm the wording with a live run: buildTutorials -live.

### Step 3

Several addresses work the same way, separated by spaces, and so does the path of a list file.

### Step 4

Every box in the dialog has an option. Dash o sends reports to a folder; dash f replaces existing reports; dash l writes the session log.

### Step 5

Dash a pauses to authenticate; dash m uses your main Edge profile; dash i runs Edge with no window.

### Step 6

So one line checks a whole list into a weekly folder, with no window shown.

Screen reader:

- urlCheck dash i dash o Weekly C colon backslash Lists backslash Site pages dot t x t

Confirm the wording with a live run: buildTutorials -live.

### Step 7

Quotation marks keep a path with spaces in one piece.

### Step 8

Saved as a batch file, that line runs on a schedule with Windows Task Scheduler, even overnight.

### Step 9

Dash u reads the settings saved by the dialog's Use configuration box, so a run tried in the dialog repeats exactly from a batch file.

### Step 10

Dash dash view output opens the folder at the end, as the dialog's box does.

### Step 11

urlCheck returns an exit code a batch file can test, so a scheduled run can flag one that failed.

Screen reader:

- if errorlevel 1 echo Check the urlCheck log

Confirm the wording with a live run: buildTutorials -live.

### Step 12

Local files can be checked by pattern: star dot h t m checks every web page in the current folder, before a site is published.

Screen reader:

- urlCheck star dot h t m dash o reports

Confirm the wording with a live run: buildTutorials -live.

### Step 13

Dash h shows every option, and dash v the version.

Screen reader:

- urlCheck dash h

Confirm the wording with a live run: buildTutorials -live.

### Step 14

Dash g opens the dialog, for anyone who starts at the prompt but prefers the fields.

### Step 15

Comparing this week's violation counts with last week's shows whether a site is getting better or worse.

### Step 16

A scheduled run can write into a folder named by date, so each week's results stay apart and comparable.

### Step 17

The session log of a scheduled run is in your application data, so a run nobody watched still leaves a record.

### Step 18

Invisible mode suits a scheduled run, since no window appears to interrupt whatever else is happening.

### Step 19

A batch file can loop over several list files, one per site, and send each to its own folder.

### Step 20

Everything learned in the dialog carries over to the prompt, option for option.

### Step 21

From the dialog, from a prompt, or on a schedule: the same checks wherever the work is. The last walk gathers it all.

**Something to try:** Write a one-line batch file that checks a list of pages into a folder, and run it.

## 9 - Conclusion

What the walks covered, a glossary in two voices, and every way to get help with urlCheck.

**Before you start:** Nothing to set up; this walk only listens.

### Step 1

That is urlCheck. Here is what each walk gave you.

### Step 2

Walk one, the dialog. Walk two, installing. Walk three, one page and its report. Walk four, the workbook and screenshots. Walk five, a list of pages.

### Step 3

Walk six, pages behind a sign in. Walk seven, the conformance report. Walk eight, the command line.

### Step 4

Now a short glossary. I say the term; the other voice says what it means.

### Step 5

axe-core.

Screen reader:

- The open-source engine urlCheck runs on each page to find accessibility problems.

### Step 6

Violation.

Screen reader:

- An element that breaks an accessibility rule, with the rule, its impact and a fix.

### Step 7

Incomplete.

Screen reader:

- A result axe could not decide by itself, left for a person to check.

### Step 8

Conformance report.

Screen reader:

- A document stating, criterion by criterion, how far a product meets WCAG; an ACR.

### Step 9

Accessibility tree.

Screen reader:

- The roles, names and states a page gives to assistive technology, as a screen reader receives them.

### Step 10: F1

Help is a key away: F1 in the dialog shows Help, Shift plus F1 says the tip for the control you are on, and F7 lists the controls.

Screen reader:

- urlCheck Help dialog

Confirm the wording with a live run: buildTutorials -live.

### Step 11: Alt+G

Guide, Alt plus G, opens the full guide and returns to the dialog. F11 checks for a newer version.

Screen reader:

- Guide button

Confirm the wording with a live run: buildTutorials -live.

### Step 12

The urlCheck page on GitHub has the latest release and a place to report a problem; each run's log says exactly what happened.

### Step 13

Thank you for listening. Happy checking.

**Something to try:** Check one page you rely on, and send its report to the site's owner.

<!-- walkthrough ends -->

## 1. Where to Go Next

Press F1 for the guide, Alt+Shift+H for the hotkey list, and Alt+F10 for every
command in one window.
