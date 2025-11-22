# __init__.py
# Copyright (C) 2025 ['Chai Chaimee & Pierre-Louis R.']
# Licensed under GNU General Public License. See COPYING.txt for details.

import wx
import api
import gui
import gui.settingsDialogs
import textInfos
import re
import controlTypes
import globalVars
from globalPluginHandler import GlobalPlugin
from scriptHandler import script
from ui import message
from NVDAObjects.IAccessible import IAccessible
from logHandler import log
import addonHandler
import json
import config
import queueHandler
import core
from enum import Enum, unique
import time
import threading
import winUser
import gui.logViewer
import os
import subprocess
import sys
import tempfile

addonHandler.initTranslation()

def fIsLogViewer(obj):
    """
    Language independent determination of a log viewer's object.
    """
    if obj is None:
        return False
    if obj.role == controlTypes.Role.PANE:
        hParent = obj.windowHandle
    else:
        hParent = winUser.getAncestor(obj.windowHandle, winUser.GA_PARENT)
    try:
        hLogViewer = gui.logViewer.logViewer.GetHandle()
        isLogViewer = hLogViewer == hParent
    except (AttributeError, RuntimeError):
        isLogViewer = False
    return isLogViewer


def initConfiguration():
    confspec = {
        "searchHistory": "string(default='[]')",
        "searchCaseSensitivity": "boolean(default=False)",
        "searchWrap": "boolean(default=True)",
        "searchType": "string(default='NORMAL')",
    }
    config.conf.spec["LogViewerPlugin"] = confspec

class SearchHistory:
    _instance = None

    @classmethod
    def get(cls):
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    def __init__(self):
        self._terms = []
        self.load()

    def load(self):
        try:
            terms = json.loads(config.conf["LogViewerPlugin"]["searchHistory"])
            if isinstance(terms, list) and all(isinstance(term, str) for term in terms):
                self._terms = terms
            else:
                log.error("Corrupted search history data, resetting to empty list.")
                self._terms = []
        except Exception as e:
            log.error(f"Error loading search history: {e}, resetting to empty list.")
            self._terms = []

    def save(self):
        try:
            config.conf["LogViewerPlugin"]["searchHistory"] = json.dumps(self._terms)
            config.conf.save()
        except Exception as e:
            log.error(f"Error saving search history: {e}")

    def getItems(self):
        return self._terms

    def getItemByText(self, text):
        return next((term for term in self._terms if term.lower() == text.lower()), None)

    def append(self, term):
        if not term:
            return
        if term.lower() in [t.lower() for t in self._terms]:
            self._terms.remove(self.getItemByText(term))
        self._terms.insert(0, term)
        if len(self._terms) > 20:
            self._terms.pop()
        self.save()

@unique
class SearchType(Enum):
    NORMAL = "normal"
    REGULAR_EXPRESSION = "regular expression"

    @staticmethod
    def getByIndex(index):
        return list(SearchType)[index]

    @staticmethod
    def getIndexByName(name):
        for index, type in enumerate(SearchType):
            if type.name == name:
                return index
        return 0  # Default to first item

    @staticmethod
    def getByName(name):
        for type in SearchType:
            if type.name == name:
                return type
        return SearchType.NORMAL  # Default to normal search

    @staticmethod
    def getSearchTypes():
        return [_(i.value) for i in SearchType]

class LogSearchDialog(wx.Dialog):
    def __init__(self, parent, logTextCtrl, globalPluginInstance):
        super().__init__(parent, title=_("Search in NVDA Log"), size=(600, 400))
        self.logCtrl = logTextCtrl
        self.dialogOpen = True
        self.searchHistory = SearchHistory.get()
        self.globalPlugin = globalPluginInstance
        self.lastSearchTerm = ""
        self.lastCaseSensitive = False
        self.lastSearchWrap = True
        self.lastSearchType = SearchType.NORMAL
        self.currentMatch = -1
        self.matches = []
        self.searchLock = threading.Lock()
        
        self.panel = wx.Panel(self)
        self.mainSizer = wx.BoxSizer(wx.VERTICAL)
        
        searchSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.searchBox = wx.ComboBox(self.panel, style=wx.CB_DROPDOWN|wx.TE_PROCESS_ENTER, choices=self.searchHistory.getItems())
        searchSizer.Add(self.searchBox, proportion=1, flag=wx.EXPAND|wx.ALL, border=5)
        
        self.searchButton = wx.Button(self.panel, label=_("Find Next"))
        searchSizer.Add(self.searchButton, flag=wx.ALIGN_CENTER_VERTICAL|wx.TOP|wx.BOTTOM|wx.RIGHT, border=5)
        
        self.findAndFocusButton = wx.Button(self.panel, label=_("Find & Focus"))
        searchSizer.Add(self.findAndFocusButton, flag=wx.ALIGN_CENTER_VERTICAL|wx.TOP|wx.BOTTOM|wx.RIGHT, border=5)
        
        self.mainSizer.Add(searchSizer, flag=wx.EXPAND)
        
        optionsSizer = wx.BoxSizer(wx.VERTICAL)
        self.caseSensitiveCheck = wx.CheckBox(self.panel, label=_("Case sensitive"))
        self.caseSensitiveCheck.SetValue(config.conf["LogViewerPlugin"]["searchCaseSensitivity"])
        optionsSizer.Add(self.caseSensitiveCheck, flag=wx.ALL, border=5)
        
        self.wrapCheck = wx.CheckBox(self.panel, label=_("Wrap around"))
        self.wrapCheck.SetValue(config.conf["LogViewerPlugin"]["searchWrap"])
        optionsSizer.Add(self.wrapCheck, flag=wx.ALL, border=5)
        
        self.searchTypeCombo = wx.Choice(self.panel, choices=SearchType.getSearchTypes())
        self.searchTypeCombo.SetSelection(SearchType.getIndexByName(config.conf["LogViewerPlugin"]["searchType"]))
        optionsSizer.Add(self.searchTypeCombo, flag=wx.ALL|wx.EXPAND, border=5)
        self.mainSizer.Add(optionsSizer, flag=wx.EXPAND)
        
        self.resultBox = wx.TextCtrl(self.panel, style=wx.TE_MULTILINE|wx.TE_READONLY)
        self.mainSizer.Add(self.resultBox, proportion=1, flag=wx.EXPAND|wx.LEFT|wx.RIGHT|wx.BOTTOM, border=5)
        
        buttonSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.prevButton = wx.Button(self.panel, label=_("Find Previous"))
        buttonSizer.Add(self.prevButton, flag=wx.ALL, border=5)
        
        self.cancelButton = wx.Button(self.panel, id=wx.ID_CANCEL, label=_("Close"))
        buttonSizer.Add(self.cancelButton, flag=wx.ALL, border=5)
        self.mainSizer.Add(buttonSizer, flag=wx.ALIGN_RIGHT|wx.ALL, border=5)
        
        self.panel.SetSizer(self.mainSizer)
        
        self.searchButton.Bind(wx.EVT_BUTTON, lambda evt: self.performSearch(forward=True, focus=False))
        self.findAndFocusButton.Bind(wx.EVT_BUTTON, lambda evt: self.performSearch(forward=True, focus=True))
        self.searchBox.Bind(wx.EVT_TEXT_ENTER, lambda evt: self.performSearch(forward=True, focus=False))
        self.prevButton.Bind(wx.EVT_BUTTON, lambda evt: self.performSearch(forward=False, focus=False))
        self.cancelButton.Bind(wx.EVT_BUTTON, self.onClose)
        self.Bind(wx.EVT_CLOSE, self.onClose)
        
        self.searchBox.SetFocus()

    def onClose(self, event):
        self.dialogOpen = False
        if self.globalPlugin and self.globalPlugin.searchDialog is self:
            self.globalPlugin.searchDialog = None
        self.Destroy()

    def doSearch(self, term, caseSensitive, searchType):
        with self.searchLock:
            self.matches = []
            try:
                textInfo = self.logCtrl.makeTextInfo(textInfos.POSITION_ALL)
                allText = textInfo.text
                
                if not allText.strip():
                    message(_("Log is empty"))
                    return False
                    
                searchFlags = 0 if caseSensitive else re.IGNORECASE
                isRegex = searchType == SearchType.REGULAR_EXPRESSION
                
                if isRegex:
                    try:
                        found_iter = re.finditer(term, allText, searchFlags)
                    except re.error as e:
                        message(_("Invalid regular expression: {error}").format(error=e))
                        return False
                else:
                    found_iter = re.finditer(re.escape(term), allText, searchFlags)

                for match in found_iter:
                    start_pos, end_pos = match.span()
                    self.matches.append((start_pos, end_pos))
                    
                self.lastSearchTerm = term
                self.lastCaseSensitive = caseSensitive
                self.lastSearchType = searchType
                
                return True
            except Exception as e:
                log.error(f"Error during search: {e}")
                return False

    def getCaretPosition(self):
        try:
            textInfo = self.logCtrl.makeTextInfo(textInfos.POSITION_CARET)
            return textInfo.bookmark.startOffset
        except Exception as e:
            log.error(f"Error getting caret position: {e}")
            return 0

    def performSearch(self, forward=True, focus=False):
        term = self.searchBox.GetValue().strip()
        if not term:
            message(_("Search term cannot be empty"))
            return
            
        self.searchHistory.append(term)
        caseSensitive = self.caseSensitiveCheck.GetValue()
        wrap = self.wrapCheck.GetValue()
        searchType = SearchType.getByIndex(self.searchTypeCombo.GetSelection())
        
        config.conf["LogViewerPlugin"]["searchCaseSensitivity"] = caseSensitive
        config.conf["LogViewerPlugin"]["searchWrap"] = wrap
        config.conf["LogViewerPlugin"]["searchType"] = searchType.name

        if (term != self.lastSearchTerm or
                caseSensitive != self.lastCaseSensitive or
                searchType != self.lastSearchType or
                not self.matches):
            if not self.doSearch(term, caseSensitive, searchType):
                self.resultBox.SetValue(_("Search failed or invalid expression"))
                message(_("No matches found"))
                return

        if not self.matches:
            self.resultBox.SetValue(_("No matches found"))
            message(_("No matches found"))
            return

        current_caret_pos = self.getCaretPosition()
        found = False
        
        if forward:
            start_index = self.currentMatch + 1 if self.currentMatch != -1 else 0
            for i in range(start_index, len(self.matches)):
                start_pos, end_pos = self.matches[i]
                if start_pos >= current_caret_pos:
                    self.currentMatch = i
                    found = True
                    break
            if not found and wrap:
                self.currentMatch = 0
                message(_("Wrapping to first match"))
                found = True
        else:
            start_index = self.currentMatch - 1 if self.currentMatch != -1 else len(self.matches) - 1
            for i in range(start_index, -1, -1):
                start_pos, end_pos = self.matches[i]
                if start_pos < current_caret_pos:
                    self.currentMatch = i
                    found = True
                    break
            if not found and wrap:
                self.currentMatch = len(self.matches) - 1
                message(_("Wrapping to last match"))
                found = True

        if not found:
            message(_("No matches found"))
            return

        self.globalPlugin.lastSearchTerm = term
        self.globalPlugin.lastMatches = self.matches
        self.globalPlugin.currentMatchIndex = self.currentMatch
        self.globalPlugin._lastSearchCaseSensitive = caseSensitive
        self.globalPlugin._lastSearchType = searchType

        self.updateResultDisplay()
        self.moveToMatch(focus)
        
        message(_("Found {count} matches.").format(count=len(self.matches)))

    def updateResultDisplay(self):
        displayText = []
        if not self.matches:
            self.resultBox.SetValue(_("No matches found"))
            return

        start_idx = max(0, self.currentMatch - 2)
        end_idx = min(len(self.matches), self.currentMatch + 3)

        textInfo = self.logCtrl.makeTextInfo(textInfos.POSITION_ALL)
        allText = textInfo.text

        for i in range(start_idx, end_idx):
            start_pos, end_pos = self.matches[i]
            
            line_num = allText.count('\n', 0, start_pos) + 1
            line_start = allText.rfind('\n', 0, start_pos) + 1
            line_end = allText.find('\n', end_pos)
            if line_end == -1:
                line_end = len(allText)
            line_text = allText[line_start:line_end].strip()

            prefix = "> " if i == self.currentMatch else "  "
            displayText.append(f"{prefix}{_('Line {number}: {text}').format(number=line_num, text=line_text)}")
            
        self.resultBox.SetValue("\n".join(displayText))
    
    def moveToMatch(self, focus=False):
        if not self.matches or self.currentMatch < 0 or self.currentMatch >= len(self.matches):
            message(_("No matches available"))
            return
            
        start_pos, end_pos = self.matches[self.currentMatch]
        try:
            if focus:
                self.Destroy()
            
            def _move():
                try:
                    focusObj = api.getFocusObject()
                    if not self.isNVDAViewerObject(focusObj):
                        if hasattr(self.logCtrl, 'setFocus'):
                            self.logCtrl.setFocus()
                        else:
                            api.setFocusObject(self.logCtrl)
                        
                    textInfo = self.logCtrl.makeTextInfo(textInfos.POSITION_ALL)
                    textInfo.collapse()
                    textInfo.move(textInfos.UNIT_CHARACTER, start_pos)
                    textInfo.collapse()
                    textInfo.updateSelection()
                    
                    line_num = textInfo.text.count('\n', 0, start_pos) + 1
                    line_start = textInfo.text.rfind('\n', 0, start_pos) + 1
                    line_end = textInfo.text.find('\n', end_pos)
                    if line_end == -1:
                        line_end = len(textInfo.text)
                    line_text = textInfo.text[line_start:line_end].strip()
                
                    message(_("Line {number}: {text}").format(number=line_num, text=line_text))
                except Exception as e:
                    log.error(f"Error moving to match: {e}")
                    message(_("Error moving to match"))
            
            queueHandler.queueFunction(queueHandler.eventQueue, _move)
        except Exception as e:
            log.error(f"Error in moveToMatch: {e}")

    def isNVDAViewerObject(self, obj):
        try:
            return obj.role == controlTypes.Role.EDITABLETEXT and fIsLogViewer(obj)
        except Exception:
            return False

class LogMonitorThread(threading.Thread):
    def __init__(self, plugin):
        super().__init__(daemon=True)
        self.plugin = plugin
        self.running = True

    def run(self):
        while self.running:
            try:
                self.plugin.backupLog()
            except Exception as e:
                log.error(f"Error in log monitor thread: {e}")
            time.sleep(5)

class GlobalPlugin(GlobalPlugin):
    bookmarkString = "BOOKMARK {0}"
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        initConfiguration()
        self.bookmarkCount = getattr(globalVars, 'devHelperBookmarkCount', 1)
        self.logViewerObj = None
        self.bookmarks = []
        self.currentBookmark = -1
        self.searchDialog = None
        self.bookmarkLock = threading.Lock()
        self.lastBookmarkRefreshTime = 0
        self.lastSearchTerm = ""
        self.lastMatches = []
        self.currentMatchIndex = -1
        self._lastSearchCaseSensitive = None
        self._lastSearchType = None
        self.logFilePointer = 0
        
        self.createOldLogFile()
        self.handleDailyReset()
        self.appendPreviousSessionLog()
        self.lastLogContent = ""
        self.initializeLogFilePointer() 
        self.backupLog(initial_backup=True)
        
        self.monitor_thread = LogMonitorThread(self)
        self.monitor_thread.start()
        
        self.isLogViewerOpen = False
    
    def initializeLogFilePointer(self):
        """Set the file pointer to the current size of nvda.log after startup operations."""
        try:
            tempDir = tempfile.gettempdir()
            logPath = os.path.join(tempDir, "nvda.log")
            if os.path.exists(logPath):
                self.logFilePointer = os.path.getsize(logPath)
            else:
                self.logFilePointer = 0
            log.debug(f"Initialized nvda.log file pointer to {self.logFilePointer} bytes.")
        except Exception as e:
            log.error(f"Error initializing log file pointer: {e}")

    def createOldLogFile(self):
        """Create oldLog.txt file in config directory if it doesn't exist"""
        try:
            configDir = globalVars.appArgs.configPath
            backupPath = os.path.join(configDir, "oldLog.txt")
            
            if not os.path.exists(backupPath):
                with open(backupPath, "w", encoding="utf-8") as f:
                    f.write("NVDA Log Backup - Created by LogViewer Plugin\n")
                    f.write("This file contains the current session's log content.\n")
                    f.write("=" * 50 + "\n\n")
                log.debug("Created oldLog.txt file")
        except Exception as e:
            log.error(f"Error creating oldLog.txt file: {e}")
    
    def handleDailyReset(self):
        """Check if the day has changed and reset oldLog.txt if necessary"""
        try:
            configDir = globalVars.appArgs.configPath
            backupPath = os.path.join(configDir, "oldLog.txt")
            if not os.path.exists(backupPath):
                return
            
            current_date = time.strftime("%Y-%m-%d")
            last_date = None
            
            with open(backupPath, "r", encoding="utf-8") as f:
                content = f.read()
                matches = list(re.finditer(r"NVDA LOG BACKUP - (\d{4}-\d{2}-\d{2})", content))
                if matches:
                    last_date = matches[-1].group(1)
            
            if last_date and last_date != current_date:
                timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
                header = f"{'='*60}\n"
                header += f"NVDA LOG BACKUP - {timestamp}\n"
                header += f"{'='*60}\n\n"
                with open(backupPath, "w", encoding="utf-8") as f:
                    f.write(header)
                log.debug("Reset oldLog.txt for new day")
        except Exception as e:
            log.error(f"Error handling daily reset: {e}")
    
    def appendPreviousSessionLog(self):
        """Append nvda-old.log to oldLog.txt if it exists"""
        try:
            tempDir = tempfile.gettempdir()
            oldNvdaLogPath = os.path.join(tempDir, "nvda-old.log")
            if not os.path.exists(oldNvdaLogPath):
                return
            
            with open(oldNvdaLogPath, "r", encoding="utf-8") as f:
                oldContent = f.read()
            
            if not oldContent.strip():
                return
            
            configDir = globalVars.appArgs.configPath
            backupPath = os.path.join(configDir, "oldLog.txt")
            
            timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
            header = f"\n\n{'='*60}\n"
            header += f"NVDA PREVIOUS SESSION LOG - {timestamp}\n"
            header += f"{'='*60}\n\n"
            
            with open(backupPath, "a", encoding="utf-8") as f:
                f.write(header + oldContent)
            
            log.debug("Appended previous session log to oldLog.txt")
        except Exception as e:
            log.error(f"Error appending previous session log: {e}")
    
    def backupLog(self, initial_backup=False):
        """
        Backup current log content to oldLog.txt - with rotation to prevent huge files.
        """
        try:
            new_content = self.getIncrementalLogContent()
            
            if not new_content.strip() and not initial_backup:
                return

            configDir = globalVars.appArgs.configPath
            backupPath = os.path.join(configDir, "oldLog.txt")
            
            if os.path.exists(backupPath):
                file_size = os.path.getsize(backupPath)
                if file_size > 5 * 1024 * 1024:
                    self.rotateOldLogFile()
            
            if new_content:
                with open(backupPath, "a", encoding="utf-8") as f:
                    if not self.lastLogContent:
                        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
                        backup_header = f"\n\n{'='*60}\n"
                        backup_header += f"NVDA LOG BACKUP - {timestamp}\n"
                        backup_header += f"{'='*60}\n\n"
                        f.write(backup_header)
                        
                    f.write(new_content)
                
                log.debug("Log content backed up to oldLog.txt")
            
            if initial_backup:
                self.lastLogContent = self.getCurrentLogContent()
                
        except Exception as e:
            log.error(f"Error backing up log: {e}")
            self.logFilePointer = 0 
    
    def rotateOldLogFile(self):
        """Rotate oldLog.txt to prevent it from growing too large"""
        try:
            configDir = globalVars.appArgs.configPath
            backupPath = os.path.join(configDir, "oldLog.txt")
            
            if not os.path.exists(backupPath):
                return
            
            with open(backupPath, "r", encoding="utf-8") as f:
                content = f.read()
            
            lines = content.split('\n')
            if len(lines) > 1000:
                recent_lines = lines[-1000:]
                
                with open(backupPath, "w", encoding="utf-8") as f:
                    f.write("NVDA Log Backup - Rotated to prevent large file size\n")
                    f.write("Previous old content discarded\n")
                    f.write("=" * 50 + "\n\n")
                    f.write('\n'.join(recent_lines))
                
                log.debug("Rotated oldLog.txt to prevent large file size. Old content discarded.")
            
        except Exception as e:
            log.error(f"Error rotating oldLog.txt: {e}")
    
    def getCurrentLogContent(self):
        """
        Get current log content by reading the whole file.
        """
        try:
            tempDir = tempfile.gettempdir()
            logPath = os.path.join(tempDir, "nvda.log")
            if os.path.exists(logPath):
                with open(logPath, "r", encoding="utf-8") as f:
                    return f.read()
            else:
                log.error("NVDA log file not found")
                return ""
        except Exception as e:
            log.error(f"Error getting current log content: {e}")
            return ""

    def getIncrementalLogContent(self):
        """
        Get only new log content since the last successful read using a file pointer.
        """
        tempDir = tempfile.gettempdir()
        logPath = os.path.join(tempDir, "nvda.log")
        if not os.path.exists(logPath):
            return ""

        try:
            current_size = os.path.getsize(logPath)
            
            if current_size < self.logFilePointer:
                log.debug("NVDA log file appears to have been cleared. Resetting file pointer.")
                self.logFilePointer = 0
                
            if current_size == self.logFilePointer:
                return ""

            with open(logPath, "r", encoding="utf-8") as f:
                f.seek(self.logFilePointer)
                new_content = f.read()
                self.logFilePointer = f.tell()
                
            return new_content
            
        except Exception as e:
            log.error(f"Error getting incremental log content: {e}")
            self.logFilePointer = 0
            return ""
    
    def monitorLogViewerState(self):
        """Monitor when log viewer is opened or closed"""
        try:
            currentState = self.isNVDAViewer()
            
            if currentState and not self.isLogViewerOpen:
                self.isLogViewerOpen = True
                log.debug("Log viewer opened - monitoring for changes")
                
            elif not currentState and self.isLogViewerOpen:
                self.isLogViewerOpen = False
                self.backupLog()
                log.debug("Log viewer closed - backed up content")
                
        except Exception as e:
            log.error(f"Error monitoring log viewer state: {e}")
    
    def terminate(self):
        """Backup log when NVDA is terminating"""
        try:
            self.monitor_thread.running = False
            self.backupLog()
            log.debug("Final log backup completed before NVDA exit")
        except Exception as e:
            log.error(f"Error during terminate backup: {e}")
        super().terminate()
    
    def isNVDAViewer(self):
        try:
            focusObj = api.getFocusObject()
            if not focusObj:
                return False
            currentState = self.isNVDAViewerObject(focusObj)
            
            if hasattr(self, 'isLogViewerOpen'):
                if currentState != self.isLogViewerOpen:
                    self.monitorLogViewerState()
            
            return currentState
        except Exception as e:
            log.error(f"Error checking NVDA Log Viewer: {e}")
            return False
    
    def isNVDAViewerObject(self, obj):
        if not obj or obj.role != controlTypes.Role.EDITABLETEXT or not fIsLogViewer(obj):
            return False
        
        self.logViewerObj = obj
        return True
    
    def isInBookmarkConflictingApp(self):
        conflicting_processes = ["notepad++", "winword", "code", "sublime_text", "atom", "brackets"]
        try:
            focusObj = api.getFocusObject()
            if not focusObj:
                return False
            processName = focusObj.appModule.appName.lower() if hasattr(focusObj.appModule, 'appName') else ""
            return processName in conflicting_processes
        except Exception as e:
            log.error(f"Error checking conflicting app: {e}")
            return False
    
    def getLogTextControl(self):
        if not self.logViewerObj:
            self.isNVDAViewer()
        return self.logViewerObj
    
    @script(description=_("Search in NVDA Log Viewer"), gesture="kb:control+f", category=_("LogViewer"))
    def script_searchInLogViewer(self, gesture):
        if not self.isNVDAViewer():
            gesture.send()
            return
            
        textCtrl = self.getLogTextControl()
        if not textCtrl:
            message(_("NVDA Log Viewer not accessible"))
            return
        
        if hasattr(self, 'searchDialog') and self.searchDialog:
            try:
                if self.searchDialog.dialogOpen:
                    self.searchDialog.Raise()
                    self.searchDialog.searchBox.SetFocus()
                    return
                else:
                    self.searchDialog.Destroy()
                    self.searchDialog = None
            except Exception:
                self.searchDialog = None
            
        def showDialog():
            try:
                app = wx.GetApp()
                if not app:
                    log.error("No wx.App instance available")
                    return
                
                topWin = app.GetTopWindow()
                if not topWin:
                    log.error("No top window available")
                    return
                
                self.searchDialog = LogSearchDialog(topWin, textCtrl, self)
                gui.mainFrame.prePopup()
                self.searchDialog.Show()
                gui.mainFrame.postPopup()
            except Exception as e:
                log.error(f"Error opening search dialog: {str(e)}")
                message(_("Failed to open search dialog"))
                
        core.callLater(100, showDialog)
    
    @script(description=_("Insert bookmark in log"), gesture="kb:control+f2", category=_("LogViewer"))
    def script_insertBookmark(self, gesture):
        if self.isInBookmarkConflictingApp():
            gesture.send()
            return
            
        bookmarkText = f"\n{self.bookmarkString.format(self.bookmarkCount)}\n"
        log.info(bookmarkText)
        message(_("Bookmark {number}").format(number=self.bookmarkCount))
        globalVars.devHelperBookmarkCount = self.bookmarkCount + 1
        self.bookmarkCount += 1
        self.backupLog()

    def _refreshBookmarks(self, textCtrl):
        current_time = time.time()
        if current_time - self.lastBookmarkRefreshTime < 0.1 and self.bookmarks:
            return
            
        with self.bookmarkLock:
            self.bookmarks = []
            if not textCtrl:
                return

            try:
                textInfo = textCtrl.makeTextInfo(textInfos.POSITION_ALL)
                all_log_text = textInfo.text
                
                if not all_log_text.strip():
                    message(_("Log is empty"))
                    return
            
                bookmark_pattern = re.compile(r"BOOKMARK (\d+)")
                
                for match in bookmark_pattern.finditer(all_log_text):
                    start_pos, end_pos = match.span()
                    bookmark_num = int(match.group(1))
                    self.bookmarks.append((start_pos, end_pos, bookmark_num))
                
                self.bookmarks.sort(key=lambda x: x[0])
                self.lastBookmarkRefreshTime = current_time

            except Exception as e:
                log.error(f"Error refreshing bookmarks: {e}")
                self.bookmarks = []

    def getCaretPosition(self, textCtrl):
        try:
            textInfo = textCtrl.makeTextInfo(textInfos.POSITION_CARET)
            return textInfo.bookmark.startOffset
        except Exception as e:
            log.error(f"Error getting caret position: {e}")
            return 0

    @script(description=_("Move to next bookmark in log"), gesture="kb:f2", category=_("LogViewer"))
    def script_moveToNextBookmark(self, gesture):
        if self.isInBookmarkConflictingApp():
            gesture.send()
            return
        
        if not self.isNVDAViewer():
            gesture.send()
            return
            
        textCtrl = self.getLogTextControl()
        if not textCtrl:
            message(_("NVDA Log Viewer not accessible"))
            return
        
        self._refreshBookmarks(textCtrl)
        
        if not self.bookmarks:
            message(_("No bookmarks found"))
            self.currentBookmark = -1
            return
            
        wrap = config.conf["LogViewerPlugin"]["searchWrap"]
        
        self.currentBookmark += 1
        if self.currentBookmark >= len(self.bookmarks):
            if wrap:
                self.currentBookmark = 0
                message(_("Wrapping to first bookmark"))
            else:
                self.currentBookmark = len(self.bookmarks) - 1
                message(_("Reached end of bookmarks"))
                return

        self._moveToBookmark(textCtrl)

    @script(description=_("Move to previous bookmark in log"), gesture="kb:shift+f2", category=_("LogViewer"))
    def script_moveToPreviousBookmark(self, gesture):
        if self.isInBookmarkConflictingApp():
            gesture.send()
            return
        
        if not self.isNVDAViewer():
            gesture.send()
            return
            
        textCtrl = self.getLogTextControl()
        if not textCtrl:
            message(_("NVDA Log Viewer not accessible"))
            return
            
        self._refreshBookmarks(textCtrl)
        
        if not self.bookmarks:
            message(_("No bookmarks found"))
            self.currentBookmark = -1
            return
            
        wrap = config.conf["LogViewerPlugin"]["searchWrap"]
        
        self.currentBookmark -= 1
        if self.currentBookmark < 0:
            if wrap:
                self.currentBookmark = len(self.bookmarks) - 1
                message(_("Wrapping to last bookmark"))
            else:
                self.currentBookmark = 0
                message(_("Already at first bookmark"))
                return
                
        self._moveToBookmark(textCtrl)
    
    def _moveToBookmark(self, textCtrl):
        if not self.bookmarks or self.currentBookmark < 0 or self.currentBookmark >= len(self.bookmarks):
            message(_("No bookmarks available"))
            return
            
        start_pos, end_pos, bookmark_num = self.bookmarks[self.currentBookmark]
        try:
            def _move():
                try:
                    focusObj = api.getFocusObject()
                    if not self.isNVDAViewerObject(focusObj):
                        if hasattr(textCtrl, 'setFocus'):
                            textCtrl.setFocus()
                        else:
                            api.setFocusObject(textCtrl)
                        
                    textInfo = textCtrl.makeTextInfo(textInfos.POSITION_ALL)
                    textInfo.collapse()
                    textInfo.move(textInfos.UNIT_CHARACTER, start_pos)
                    textInfo.collapse()
                    textInfo.updateSelection()
                    message(_("Bookmark {number}").format(number=bookmark_num))
                except Exception as e:
                    log.error(f"Error moving to bookmark: {e}")
                    message(_("Error moving to bookmark"))
            
            queueHandler.queueFunction(queueHandler.eventQueue, _move)
        except Exception as e:
            log.error(f"Error in _moveToBookmark: {e}")

    def _doQuickSearch(self, term, caseSensitive, searchType):
        try:
            textCtrl = self.getLogTextControl()
            if not textCtrl:
                return False
                
            textInfo = textCtrl.makeTextInfo(textInfos.POSITION_ALL)
            allText = textInfo.text
            if not allText.strip():
                message(_("Log is empty"))
                return False
                
            searchFlags = 0 if caseSensitive else re.IGNORECASE
            isRegex = searchType == SearchType.REGULAR_EXPRESSION
            
            if isRegex:
                try:
                    self.lastMatches = [m.span() for m in re.finditer(term, allText, searchFlags)]
                except re.error as e:
                    message(_("Invalid regular expression: {error}").format(error=e))
                    return False
            else:
                self.lastMatches = [m.span() for m in re.finditer(re.escape(term), allText, searchFlags)]
            
            self.lastSearchTerm = term
            self._lastSearchCaseSensitive = caseSensitive
            self._lastSearchType = searchType
            return True
        except Exception as e:
            log.error(f"Error during quick search: {e}")
            return False

    @script(description=_("Find next occurrence using last search term"), gesture="kb:f3", category=_("LogViewer"))
    def script_findNext(self, gesture):
        if not self.isNVDAViewer():
            gesture.send()
            return
            
        textCtrl = self.getLogTextControl()
        if not textCtrl:
            message(_("NVDA Log Viewer not accessible"))
            return
            
        if not self.lastSearchTerm:
            message(_("No search has been performed yet. Please use Control+F to perform a search first."))
            return

        caseSensitive = config.conf["LogViewerPlugin"]["searchCaseSensitivity"]
        searchType = SearchType.getByName(config.conf["LogViewerPlugin"]["searchType"])
        wrap = config.conf["LogViewerPlugin"]["searchWrap"]
        
        if (
            not self.lastMatches or
            caseSensitive != self._lastSearchCaseSensitive or
            searchType != self._lastSearchType
        ):
            if not self._doQuickSearch(self.lastSearchTerm, caseSensitive, searchType):
                message(_("No matches found"))
                return
            
        if not self.lastMatches:
            message(_("No matches found"))
            return

        self.currentMatchIndex += 1
        if self.currentMatchIndex >= len(self.lastMatches):
            if wrap:
                self.currentMatchIndex = 0
                message(_("Wrapping to first match"))
            else:
                self.currentMatchIndex = len(self.lastMatches) - 1
                message(_("Reached end of matches"))
                return

        self._moveToQuickSearchResult(textCtrl)
    
    @script(description=_("Find previous occurrence using last search term"), gesture="kb:shift+f3", category=_("LogViewer"))
    def script_findPrevious(self, gesture):
        if not self.isNVDAViewer():
            gesture.send()
            return
            
        textCtrl = self.getLogTextControl()
        if not textCtrl:
            message(_("NVDA Log Viewer not accessible"))
            return
            
        if not self.lastSearchTerm:
            message(_("No search has been performed yet. Please use Control+F to perform a search first."))
            return

        caseSensitive = config.conf["LogViewerPlugin"]["searchCaseSensitivity"]
        searchType = SearchType.getByName(config.conf["LogViewerPlugin"]["searchType"])
        wrap = config.conf["LogViewerPlugin"]["searchWrap"]
        
        if (
            not self.lastMatches or
            caseSensitive != self._lastSearchCaseSensitive or
            searchType != self._lastSearchType
        ):
            if not self._doQuickSearch(self.lastSearchTerm, caseSensitive, searchType):
                message(_("No matches found"))
                return
                
        if not self.lastMatches:
            message(_("No matches found"))
            return

        self.currentMatchIndex -= 1
        if self.currentMatchIndex < 0:
            if wrap:
                self.currentMatchIndex = len(self.lastMatches) - 1
                message(_("Wrapping to last match"))
            else:
                self.currentMatchIndex = 0
                message(_("Already at first match"))
                return
                
        self._moveToQuickSearchResult(textCtrl)
    
    def _moveToQuickSearchResult(self, textCtrl):
        if not self.lastMatches or self.currentMatchIndex < 0 or self.currentMatchIndex >= len(self.lastMatches):
            message(_("No matches available"))
            return
            
        start_pos, end_pos = self.lastMatches[self.currentMatchIndex]
        try:
            def _move():
                try:
                    focusObj = api.getFocusObject()
                    if not self.isNVDAViewerObject(focusObj):
                        if hasattr(textCtrl, 'setFocus'):
                            textCtrl.setFocus()
                        else:
                            api.setFocusObject(textCtrl)

                    textInfo = textCtrl.makeTextInfo(textInfos.POSITION_ALL)
                    textInfo.collapse()
                    textInfo.move(textInfos.UNIT_CHARACTER, start_pos)
                    textInfo.collapse()
                    textInfo.updateSelection()
                    
                    message(_("{term} {current} of {total}").format(term=self.lastSearchTerm, current=self.currentMatchIndex + 1, total=len(self.lastMatches)))
                except Exception as e:
                    log.error(f"Error moving to quick search result: {e}")
                    message(_("Error moving to match"))

            queueHandler.queueFunction(queueHandler.eventQueue, _move)
        except Exception as e:
            log.error(f"Error in _moveToQuickSearchResult: {e}")

    @script(description=_("Open old log file in default editor"), gesture="kb:NVDA+control+l", category=_("LogViewer"))
    def script_openOldLog(self, gesture):
        try:
            configDir = globalVars.appArgs.configPath
            backupPath = os.path.join(configDir, "oldLog.txt")
            
            if not os.path.exists(backupPath):
                message(_("Old log file not found"))
                return
            
            # Use subprocess to open the file with the default associated program
            if sys.platform.startswith("win"):
                os.startfile(backupPath)
            else:
                subprocess.run(["xdg-open", backupPath], check=True)
            message(_("Opening old log file"))
        except Exception as e:
            log.error(f"Error opening old log file: {e}")
            message(_("Failed to open old log file"))
