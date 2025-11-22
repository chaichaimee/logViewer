<div align="center">

# logViewer

**Authors:** chai chaimee & Pierre-Louis R.  
**URL:** https://github.com/chaichaimee/logViewer

</div>

## What's New

Added **oldLog** function – a powerful way to preserve and access logs even after NVDA crashes and restarts.

## Features

logViewer significantly enhances NVDA's built-in Log Viewer, making it much more powerful for developers and advanced users who need to analyze logs efficiently. Key features include:

- **Advanced Search:**  
  Press **Control+F** to open the search dialog.  
  Supports case-sensitive search, wrap-around, regular expressions, and search history for quick reuse of previous terms.

- **Quick Search Navigation:**  
  Jump to the next search result with **F3**, or go back to the previous one with **Shift+F3** — without reopening the search dialog.

- **Bookmark System:**  
  Add bookmarks using **Control+F2**.  
  Navigate between bookmarks using **F2** (next) and **Shift+F2** (previous).

- **Old Log Backup & Access:**  
  Automatically maintains an `oldLog.txt` file containing logs from the current and previous sessions (including after crashes).  
  Open the preserved log anytime with **NVDA+Control+L**.

## Old Log Functionality

The logViewer add-on includes a robust system to manage and preserve NVDA log content in an `oldLog.txt` file, ensuring logs from current and previous sessions are backed up and easily accessible — especially useful after crashes.

### Key Features

- **Automatic Log Backup**: Continuously saves new NVDA log content to `oldLog.txt` in the NVDA configuration directory every 5 seconds, using efficient incremental reading.

- **Previous Session Logs**: When NVDA starts, the previous session’s log (`nvda-old.log`) is automatically appended to `oldLog.txt` with a clear timestamped header.

- **Daily Reset**: If the date has changed since the last run, `oldLog.txt` starts fresh with a new timestamped header, keeping logs organized by day.

- **File Rotation**: When `oldLog.txt` exceeds 5 MB, the oldest content is archived into a dated backup file (e.g., `oldLog_20251122_143022.txt`), and the last 1,000 lines are retained.

- **Quick Access**: Press **NVDA+Ctrl+L** to instantly open `oldLog.txt` in your default text editor.

### Handling NVDA Crash and Restart

When NVDA crashes and automatically restarts:

- The previous session’s log (`nvda-old.log`) is automatically appended to `oldLog.txt` with a timestamped header during add-on initialization.

- The add-on detects if the current `nvda.log` is smaller than expected (indicating a crash/restart) and correctly resets its reading position.

- Continuous incremental backup resumes immediately, ensuring no log data is lost.

This guarantees that even after unexpected crashes, all relevant log information is safely preserved and ready for debugging.

## Language Support

This add-on works seamlessly with NVDA in **any language**.

Instead of relying on window titles (which vary by language), logViewer identifies the NVDA Log Viewer by its window handle.  
This approach ensures full compatibility across all languages supported by NVDA.
