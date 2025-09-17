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

addonHandler.initTranslation()

def fIsLogViewer(obj):
    """
    language independant  determination of a log viewer's object.
    """
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
    
    def isNVDAViewer(self):
        try:
            focusObj = api.getFocusObject()
            return focusObj and self.isNVDAViewerObject(focusObj)
        except Exception as e:
            log.error(f"Error checking NVDA Log Viewer: {e}")
        return False
    
    def isNVDAViewerObject(self, obj):
        if obj.role != controlTypes.Role.EDITABLETEXT or not fIsLogViewer(obj):
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
        
        if hasattr(self, 'searchDialog') and self.searchDialog and self.searchDialog.dialogOpen:
            try:
                self.searchDialog.Raise()
                self.searchDialog.searchBox.SetFocus()
                return
            except Exception:
                if self.searchDialog:
                    self.searchDialog.Destroy()
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

    def isOnBookmark(self, textCtrl):
        try:
            caret_pos = self.getCaretPosition(textCtrl)
            for bookmark in self.bookmarks:
                start_pos, end_pos, _ = bookmark
                if start_pos <= caret_pos <= end_pos:
                    return True
            return False
        except Exception as e:
            log.error(f"Error checking if on bookmark: {e}")
            return False

    @script(description=_("Jump to next bookmark in log"), gesture="kb:f2", category=_("LogViewer"))
    def script_jumpToNextBookmark(self, gesture):
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
            
        current_caret_pos = self.getCaretPosition(textCtrl)
        found_next = False
        wrap = config.conf["LogViewerPlugin"]["searchWrap"]
        
        if self.isOnBookmark(textCtrl):
            current_bookmark_end = -1
            for bookmark in self.bookmarks:
                start_pos, end_pos, _ = bookmark
                if start_pos <= current_caret_pos <= end_pos:
                    current_bookmark_end = end_pos
                    break
            
            if current_bookmark_end != -1:
                for i, bookmark in enumerate(self.bookmarks):
                    start_pos, _, _ = bookmark
                    if start_pos > current_bookmark_end:
                        self.currentBookmark = i
                        found_next = True
                        break
        
        if not found_next:
            for i, bookmark in enumerate(self.bookmarks):
                start_pos, end_pos, _ = bookmark
                if start_pos > current_caret_pos:
                    self.currentBookmark = i
                    found_next = True
                    break
        
        if not found_next:
            if wrap:
                self.currentBookmark = 0
                message(_("Wrapping to first bookmark"))
                found_next = True
            else:
                message(_("Reached end of bookmarks"))
                return

        if found_next:
            self._moveToBookmark(textCtrl)

    @script(description=_("Jump to previous bookmark in log"), gesture="kb:shift+f2", category=_("LogViewer"))
    def script_jumpToPreviousBookmark(self, gesture):
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
            
        current_caret_pos = self.getCaretPosition(textCtrl)
        found_prev = False
        wrap = config.conf["LogViewerPlugin"]["searchWrap"]
        
        if self.isOnBookmark(textCtrl):
            current_bookmark_start = -1
            for bookmark in self.bookmarks:
                start_pos, end_pos, _ = bookmark
                if start_pos <= current_caret_pos <= end_pos:
                    current_bookmark_start = start_pos
                    break

            if current_bookmark_start != -1:
                for i in range(len(self.bookmarks) - 1, -1, -1):
                    start_pos, _, _ = self.bookmarks[i]
                    if start_pos < current_bookmark_start:
                        self.currentBookmark = i
                        found_prev = True
                        break

        if not found_prev:
            for i in range(len(self.bookmarks) - 1, -1, -1):
                start_pos, end_pos, _ = self.bookmarks[i]
                if start_pos < current_caret_pos:
                    self.currentBookmark = i
                    found_prev = True
                    break

        if not found_prev:
            if wrap:
                self.currentBookmark = len(self.bookmarks) - 1
                message(_("Wrapping to last bookmark"))
                found_prev = True
            else:
                message(_("Already at first bookmark"))
                return
                
        if found_prev:
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

    @script(description=_("Find next match"), gesture="kb:f3", category=_("LogViewer"))
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
    
    @script(description=_("Find previous match"), gesture="kb:shift+f3", category=_("LogViewer"))
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
                    
                    # Use the actual search term in the message instead of "Match"
                    message(_("{term} {current} of {total}").format(term=self.lastSearchTerm, current=self.currentMatchIndex + 1, total=len(self.lastMatches)))
                except Exception as e:
                    log.error(f"Error moving to quick search result: {e}")
                    message(_("Error moving to match"))

            queueHandler.queueFunction(queueHandler.eventQueue, _move)
        except Exception as e:
            log.error(f"Error in _moveToQuickSearchResult: {e}")
