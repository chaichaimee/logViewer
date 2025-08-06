# logViewer

**Add-on Name:** logViewer  
**Summary:** Enhance the NVDA Log Viewer with advanced search and bookmarking features for easier function tracking and debugging.  
**Authors:** chai chaimee & Pierre-Louis R.  
**URL:** [https://github.com/chaichaimee/logViewer](https://github.com/chaichaimee/logViewer)  

## Features

logViewer improves the functionality of NVDA's built-in Log Viewer, making it more powerful for developers and advanced users who need to analyze logs efficiently. Key features include:  

- **Advanced Search:**  
  Press `Control+F` to open the search dialog.  
  Supports case-sensitive search, wrap-around, regular expressions, and search history for quick reuse of previous terms.  

- **Quick Search Navigation:**  
  Jump to the next search result with `F3`, or go back to the previous one with `Shift+F3` — without reopening the search dialog.  

- **Bookmark System:**  
  Add bookmarks using `Control+F2`.  
  Navigate between bookmarks using `F2` (next) and `Shift+F2` (previous).  

## Language Support

This add-on is designed to work seamlessly with NVDA in any language.  
Instead of checking the window title (which depends on UI language), it detects the NVDA Log Viewer using its window handle.  
This ensures full compatibility across all languages NVDA supports.  
