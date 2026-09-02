<p align="center">
  <img src="https://www.nvaccess.org/files/nvda/documentation/userGuide/images/nvda.ico" alt="NVDA Logo" width="120">
</p>

<h1 align="center">LogViewer</h1>

<p align="center">Find it, mark it, copy it — total command of the NVDA log.</p>

<p align="center">
  <b>author:</b> chai chaimee & Pierre-Louis R.<br>
  <b>url:</b> <a href="https://github.com/chaichaimee/logViewer">https://github.com/chaichaimee/logViewer</a>
</p>

## Introduction

LogViewer is an NVDA add-on that turns the NVDA Log Viewer, and even a log file opened in an external text editor, into a fully searchable and bookmarkable workspace.

Instead of scrolling endlessly through thousands of log lines to find an error, you can search instantly, jump straight to any bookmark, and copy an entire error block to the clipboard with a single extra tap of a key.

It also watches for crashes automatically: if NVDA restarted after a crash, LogViewer quietly drops a bookmark at the point of failure in the old log so you can find it right away.

## Hot Keys

> **Control+F**  
> Single Tap : Open the Search in NVDA Log dialog (only works while focused in the NVDA Log Viewer)

> **Control+F2**  
> Single Tap : Insert a numbered bookmark at the current point in the log

> **F2**  
> Single Tap : Move to the next bookmark in the log

> **Shift+F2**  
> Single Tap : Move to the previous bookmark in the log

> **F3**  
> Single Tap : Find the next occurrence of the current quick-search term  
> Double Tap : Copy the ERROR / WARNING / Traceback block at the current match to the clipboard

> **Shift+F3**  
> Single Tap : Find the previous occurrence of the current quick-search term

> **NVDA+Control+L**  
> Single Tap : Open the NVDA log file in your default text editor (prefers nvda-old.log, falls back to nvda.log)

> **Control+Shift+F2**  
> Single Tap : Copy everything logged since the latest bookmark to the clipboard  
> Double Tap : Step back and copy the segment belonging to the previous bookmark  
> Triple Tap : Switch between reading the current log and the old log for this feature

## Features

### Search Dialog (Control+F)

While focused in the NVDA Log Viewer, Control+F opens a dedicated search dialog with:

* A search box that remembers your last 20 search terms.
* A **Case sensitive** checkbox.
* A **Wrap around** checkbox, so searching continues from the top/bottom instead of stopping at the end.
* A choice between **normal text** and **regular expression** search.
* **Find Next** and **Find Previous** buttons to move through matches while the dialog stays open.
* A **Find & Focus** button (or pressing Enter in the search box) that closes the dialog and moves your cursor directly onto the match inside the log.
* A results box that previews a few matches around your current position, marking the active one.

The search itself runs on a background thread so that very large logs do not slow down or freeze NVDA while scanning.

### Bookmarks (Control+F2, F2, Shift+F2)

Control+F2 writes a line such as **BOOKMARK 3** into the log at your current position and speaks the bookmark number. Each new bookmark automatically increases the number by one.

F2 and Shift+F2 then let you jump forward or backward between these bookmarks. If **Wrap around** is enabled in settings, reaching the last bookmark loops back to the first, and vice versa.

Step by step, when you press F2:

1. LogViewer re-scans the log text for all **BOOKMARK** markers and sorts them by position.
2. It advances one position forward in that sorted list.
3. If you were already on the last bookmark and wrap is off, you hear "Reached end of bookmarks" and stay put.
4. Otherwise your cursor is moved to that bookmark's location and NVDA announces the bookmark number.

The bookmark counter itself resets to 1 automatically whenever Windows has been fully rebooted (detected via system uptime), but stays the same across ordinary NVDA restarts.

### Automatic Crash Bookmark

On startup, LogViewer checks the previous session's log (nvda-old.log) for signs that NVDA crashed, such as a Traceback, an unhandled exception, or a CRASH entry as the very last line.

If a crash is detected and a bookmark was not already placed near it, LogViewer silently appends a bookmark to that old log so you can locate the crash point straight away with F2/Shift+F2, without having to manually search for it.

### Quick Search and Copy Error Block (F3 / Shift+F3)

F3 performs a quick, dialog-free search using the term last searched, or, if none has been used yet, one of your configured default quick-search terms (see Settings below).

F3 is also a double-action key:

* **Single tap**: jump to the next match and announce it (the exact announcement format depends on your settings).
* **Double tap** (within half a second): instead of moving, LogViewer copies the whole error/warning/traceback block surrounding the current match to the clipboard, and plays a short beep to confirm the copy.

How the "block" is detected, step by step:

1. LogViewer looks upward from the match to find where the block starts, recognizing lines that begin with "ERROR -", "WARNING -", or "Traceback".
2. It then looks downward, including any indented continuation lines that belong to that same block (for example, the lines of a Python traceback).
3. It stops at the next INFO/WARNING/ERROR/DEBUG line, a blank line, or an unindented line, whichever comes first.
4. Only the resulting block of text is copied — not the entire log.

Shift+F3 always performs a plain "find previous" using the same last search term and never triggers the copy behavior.

### Open Log File (NVDA+Control+L)

This opens the NVDA log in your system's default text editor. It prefers nvda-old.log (the previous session's log, most useful right after a crash or restart) and falls back to nvda.log if no old log exists.

Once opened this way, LogViewer remembers which file you opened so that F2, Shift+F2, and the search hotkeys can keep working even while your cursor is inside that external editor window, not just inside the built-in NVDA Log Viewer.

### Copy From Bookmark (Control+Shift+F2)

This hotkey reads bookmarks directly from the log file on disk, so it works even if neither the NVDA Log Viewer nor an external editor is open. It has three tap levels:

* **Single tap**: copies everything logged after the most recent bookmark to the clipboard.
* **Double tap**: steps one bookmark further back and copies that earlier segment instead — pressing it repeatedly walks backward through your bookmark history, looping back to the latest bookmark once you pass the first.
* **Triple tap**: switches which log file this feature reads from, toggling between the current session's nvda.log and the previous session's nvda-old.log.

Each successful copy plays a short confirmation beep.

### Conflicting Application Awareness

Some applications use Control+F2 or Control+Shift+F2 for their own purposes. LogViewer checks the current application (for example Notepad++, Word, VS Code, Sublime Text, Atom, or Brackets) and, if it matches one of these, passes the key press straight through to that application instead of intercepting it.

### Settings Panel

A "Log Viewer" category is added to the NVDA Settings dialog, offering:

* **Case sensitive by default**, and **Wrap around by default**, checkboxes matching the search dialog's own options.
* **Search result announcement style**, choosing whether results are read as term, line number and match position together; line number only; or just the matched line's text.
* **Default quick-search terms**, a list (one per line) of terms that F3 falls back to automatically when no manual search has been made yet.

All settings, along with your search term history, are stored as JSON files under your NVDA user configuration folder, and LogViewer automatically migrates data saved by older versions of the add-on to this location.

## Support Me

If this tool has made your life easier, consider fueling the next update with a small donation.

[![Support me](https://img.shields.io/badge/Donate-Support%20Me-blue?style=for-the-badge&logo=stripe)](https://buy.stripe.com/dRm9AU1xQ3Ds22N6VK1VK01)

Your support means the world. Let's build something great together

<p align="center">
  <sub>© 2026 Chai Chaimee NVDA Add-on Released under GNU GPL</sub>
</p>