# addonCore.py

import wx
import api
import gui
import gui.settingsDialogs
import gui.logViewer
import textInfos
import controlTypes
import core
import config
import ctypes
import os
import threading
import time
import tones
import weakref
import addonHandler
from collections import namedtuple
from globalPluginHandler import GlobalPlugin as BaseGlobalPlugin
from scriptHandler import script
from ui import message
from NVDAObjects.window import Window as NVDAWindowObject
from logHandler import log

from .bookmark_logic import countLines, extractSegment, findBookmarks, pickCursorIndex, pickLogPath
from .config_manager import (
	PluginSettings,
	SearchHistory,
	importLegacyNvdaIniSettings,
	initConfiguration,
)
from .log_files import getCurrentLogPath, getOldLogPath, readLogText
from .search_dialog import LogSearchDialog
from .search_logic import (
	SearchManager,
	SearchType,
	describeSearchFailure,
	fIsLogViewer,
	get_block_at_position,
	get_full_text,
)
from .settings_panel import LogViewerSettingsPanel

addonHandler.initTranslation()

BookmarkCopyResult = namedtuple(
	"BookmarkCopyResult",
	("status", "text", "bookmarkNumber", "cursorIndex", "wrapped", "logPath"),
)


def buildBookmarkCopy(preferOld, cursor, stepBack):
	"""Runs on a worker thread. Reads the log once and derives bookmarks and segment from that one snapshot."""
	logPath = pickLogPath(preferOld, getCurrentLogPath(), getOldLogPath(), os.path.isfile)
	if not logPath:
		return BookmarkCopyResult("noLog", "", 0, None, False, None)
	logText = readLogText(logPath)
	if logText is None:
		return BookmarkCopyResult("readFailed", "", 0, None, False, logPath)
	bookmarks = findBookmarks(logText)
	cursorIndex, wrapped = pickCursorIndex(len(bookmarks), cursor, stepBack)
	if cursorIndex is None:
		return BookmarkCopyResult("wholeLog", logText.strip("\r\n"), 0, None, False, logPath)
	segmentText = extractSegment(logText, bookmarks, cursorIndex)
	return BookmarkCopyResult("ok", segmentText, bookmarks[cursorIndex][2], cursorIndex, wrapped, logPath)


def loadBookmarksFromFile(filePath):
	logText = readLogText(filePath)
	if logText is None:
		return []
	return findBookmarks(logText)


class GlobalPlugin(BaseGlobalPlugin):
	bookmarkString = "BOOKMARK {0}"

	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
		self._isTerminated = False
		initConfiguration()
		SearchHistory.get()
		PluginSettings.get()
		importLegacyNvdaIniSettings()
		PluginSettings.get().ensureFileExists()
		self.bookmarkCount = config.conf["LogViewerPlugin"].get("bookmarkCount", 1)
		self._resetBookmarkCountOnNewBoot()
		self._logViewerWeakRef = None
		self.bookmarks = []
		self.currentBookmark = -1
		self.bookmarkLock = threading.Lock()
		self.lastBookmarkRefreshTime = 0
		self.current_log_file = None
		self._searchDialogOpen = False
		threading.Thread(target=self._addCrashBookmarkIfNeeded, daemon=True).start()

		self.search_manager = SearchManager()

		self._findNext_tap_time = 0
		self._findNext_tap_count = 0
		self._tap_threshold = 0.5
		self._findNext_tap_timer = None

		# Ctrl+Shift+F2 state: cursor is an index into the sorted bookmark list of the log the
		# gesture reads from; None means the double-tap has not stepped back yet.
		self._copyBookmark_tap_time = 0
		self._copyBookmark_tap_count = 0
		self._copyBookmark_tap_timer = None
		self._clipboardBookmarkCursor = None
		# The current session's nvda.log is where new bookmarks land, so it is the default;
		# nvda-old.log is only read after a triple-tap switch.
		self._clipboardLogMode = "current"

		if LogViewerSettingsPanel not in gui.settingsDialogs.NVDASettingsDialog.categoryClasses:
			gui.settingsDialogs.NVDASettingsDialog.categoryClasses.append(LogViewerSettingsPanel)

	def _getSystemBootTimestamp(self):
		try:
			uptimeMs = ctypes.windll.kernel32.GetTickCount64()
			return time.time() - (uptimeMs / 1000.0)
		except OSError as e:
			log.error(f"Error reading system uptime for boot detection: {e}")
			return None

	def _resetBookmarkCountOnNewBoot(self):
		currentBootTimestamp = self._getSystemBootTimestamp()
		if currentBootTimestamp is None:
			return
		storedBootTimestamp = config.conf["LogViewerPlugin"].get("lastBootTimestamp", 0.0)
		# A few seconds of slack absorbs clock rounding between GetTickCount64 reads of the same boot.
		if abs(currentBootTimestamp - float(storedBootTimestamp)) > 5.0:
			self.bookmarkCount = 1
			config.conf["LogViewerPlugin"]["bookmarkCount"] = 1
			config.conf["LogViewerPlugin"]["lastBootTimestamp"] = currentBootTimestamp
			config.conf.save()

	def terminate(self):
		self._isTerminated = True
		for timerName in ("_findNext_tap_timer", "_copyBookmark_tap_timer"):
			pendingTimer = getattr(self, timerName, None)
			if pendingTimer:
				pendingTimer.Stop()
				setattr(self, timerName, None)
		self.search_manager.invalidate()
		self.bookmarks = []
		self._logViewerWeakRef = None
		try:
			if LogViewerSettingsPanel in gui.settingsDialogs.NVDASettingsDialog.categoryClasses:
				gui.settingsDialogs.NVDASettingsDialog.categoryClasses.remove(LogViewerSettingsPanel)
		except (AttributeError, ValueError, RuntimeError) as e:
			log.error(f"Error during terminate: {e}")
		super().terminate()

	def _addCrashBookmarkIfNeeded(self):
		# Worker thread: file work only. The counter and config.conf belong to the main thread
		# and are updated through wx.CallAfter in _commitCrashBookmarkCount.
		try:
			oldLogPath = getOldLogPath()
			if not os.path.isfile(oldLogPath):
				return
			with open(oldLogPath, 'r', encoding='utf-8', errors='replace') as oldLogFile:
				lines = oldLogFile.readlines()
			if not lines:
				return
			lastLine = lines[-1].strip()
			crashIndicators = ["Traceback", "ERROR - unhandled exception", "CRASH"]
			if not any(indicator in lastLine for indicator in crashIndicators):
				return
			if any(line.strip().startswith("BOOKMARK") for line in lines[-5:]):
				return
			pendingBookmarkNumber = self.bookmarkCount
			bookmarkLine = f"\n{self.bookmarkString.format(pendingBookmarkNumber)}\n"
			with open(oldLogPath, 'a', encoding='utf-8') as oldLogFile:
				oldLogFile.write(bookmarkLine)
			log.info(f"Added crash bookmark number {pendingBookmarkNumber} to {oldLogPath}")
			wx.CallAfter(self._commitCrashBookmarkCount, pendingBookmarkNumber + 1)
		except OSError as e:
			log.error(f"Error adding crash bookmark: {e}")

	def _commitCrashBookmarkCount(self, newCount):
		if self._isTerminated:
			return
		self.bookmarkCount = newCount
		config.conf["LogViewerPlugin"]["bookmarkCount"] = newCount
		config.conf.save()

	def isNVDAViewer(self):
		try:
			focusObj = api.getFocusObject()
			if not focusObj:
				return False
			return self.isNVDAViewerObject(focusObj)
		except (AttributeError, RuntimeError) as e:
			log.error(f"Error checking NVDA Log Viewer: {e}")
			return False

	def isNVDAViewerObject(self, obj):
		if not obj or obj.role != controlTypes.Role.EDITABLETEXT or not fIsLogViewer(obj):
			return False
		self._logViewerWeakRef = weakref.ref(obj)
		return True

	def isInBookmarkConflictingApp(self):
		conflicting_processes = ["notepad++", "winword", "code", "sublime_text", "atom", "brackets"]
		try:
			focusObj = api.getFocusObject()
			if not focusObj:
				return False
			processName = focusObj.appModule.appName.lower() if hasattr(focusObj.appModule, 'appName') else ""
			return processName in conflicting_processes
		except AttributeError as e:
			log.error(f"Error checking conflicting app: {e}")
			return False

	def getLogTextControl(self):
		if self._logViewerWeakRef:
			obj = self._logViewerWeakRef()
			if obj and self.isNVDAViewerObject(obj):
				return obj
		if self.isNVDAViewer():
			return self._logViewerWeakRef() if self._logViewerWeakRef else None
		return None

	def _isExternalLogEditor(self, obj):
		if not obj or obj.role != controlTypes.Role.EDITABLETEXT:
			return False
		if not self.current_log_file:
			return False
		try:
			window_title = obj.windowText or ""
			base_name = os.path.basename(self.current_log_file)
			return base_name.lower() in window_title.lower()
		except AttributeError:
			return False

	def _getExternalLogTextControl(self):
		focus = api.getFocusObject()
		if self._isExternalLogEditor(focus):
			return focus
		return None

	def _copyTextToClipboard(self, text):
		if not wx.TheClipboard.Open():
			log.error("Could not open clipboard")
			return False
		try:
			wx.TheClipboard.SetData(wx.TextDataObject(text))
		finally:
			wx.TheClipboard.Close()
		return True

	def _scheduleConfirmationBeep(self, delay):
		def playBeep():
			try:
				tones.beep(440, 100)
			except RuntimeError as e:
				log.error(f"Error during beep: {e}")
		core.callLater(delay, playBeep)

	@script(description=_("Search in NVDA Log Viewer"), gesture="kb:control+f", category=_("LogViewer"))
	def script_searchInLogViewer(self, gesture):
		if not self.isNVDAViewer():
			gesture.send()
			return
		textCtrl = self.getLogTextControl()
		if not textCtrl:
			message(_("NVDA Log Viewer not accessible"))
			return
		if self._searchDialogOpen:
			return
		self._searchDialogOpen = True
		# The dialog is modal but must not block this script.
		wx.CallAfter(self._showSearchDialog, textCtrl)

	def _getSearchDialogParent(self):
		viewerFrame = getattr(gui.logViewer, "logViewer", None)
		return viewerFrame if viewerFrame is not None else gui.mainFrame

	def _showSearchDialog(self, textCtrl):
		request = None
		gui.mainFrame.prePopup()
		try:
			dialog = LogSearchDialog(self._getSearchDialogParent())
			try:
				if dialog.ShowModal() == wx.ID_OK:
					request = dialog.request
			finally:
				dialog.Destroy()
		except RuntimeError as e:
			log.error(f"Error opening search dialog: {e}")
			message(_("Failed to open search dialog"))
		finally:
			gui.mainFrame.postPopup()
			self._searchDialogOpen = False
		if request is None or self._isTerminated:
			return
		# Give the log viewer time to take focus back before the caret moves.
		core.callLater(100, self._startSearch, textCtrl, request)

	def _startSearch(self, textCtrl, request):
		if self._isTerminated:
			return
		manager = self.search_manager
		isNewSearch = not manager.matchesQuery(request.term, request.caseSensitive, request.searchType)

		def onReady(status):
			if self._isTerminated:
				return
			if status != "ok":
				message(describeSearchFailure(status))
				return
			self._stepToMatch(textCtrl, True, request.wrap, inclusive=isNewSearch)

		manager.requestMatches(textCtrl, request.term, request.caseSensitive, request.searchType, onReady)

	@script(description=_("Insert bookmark in log"), gesture="kb:control+f2", category=_("LogViewer"))
	def script_insertBookmark(self, gesture):
		if self.isInBookmarkConflictingApp():
			gesture.send()
			return
		bookmarkText = f"\n{self.bookmarkString.format(self.bookmarkCount)}\n"
		log.info(bookmarkText)
		message(_("Bookmark {number}").format(number=self.bookmarkCount))
		self.bookmarkCount += 1
		config.conf["LogViewerPlugin"]["bookmarkCount"] = self.bookmarkCount
		config.conf.save()

	def _refreshBookmarks(self, textCtrl, onComplete=None):
		# The control is read here on the main thread; only the plain string goes to the worker.
		currentTime = time.time()
		if currentTime - self.lastBookmarkRefreshTime < 0.1 and self.bookmarks:
			if onComplete:
				onComplete()
			return

		allLogText = get_full_text(textCtrl)
		if not allLogText or not allLogText.strip():
			message(_("Log is empty"))
			self._applyRefreshedBookmarks([], onComplete)
			return

		def worker():
			foundBookmarks = findBookmarks(allLogText)
			wx.CallAfter(self._applyRefreshedBookmarks, foundBookmarks, onComplete)

		threading.Thread(target=worker, daemon=True).start()

	def _applyRefreshedBookmarks(self, foundBookmarks, onComplete):
		if self._isTerminated:
			return
		with self.bookmarkLock:
			self.bookmarks = foundBookmarks
			self.lastBookmarkRefreshTime = time.time()
		if onComplete:
			onComplete()

	def getCaretPosition(self, textCtrl):
		try:
			textInfo = textCtrl.makeTextInfo(textInfos.POSITION_CARET)
			return textInfo.bookmark.startOffset
		except (RuntimeError, NotImplementedError) as e:
			log.error(f"Error getting caret position: {e}")
			return 0

	def _navigateBookmark(self, gesture, forward):
		if self.isInBookmarkConflictingApp():
			gesture.send()
			return

		if self.isNVDAViewer():
			textCtrl = self.getLogTextControl()
			if not textCtrl:
				message(_("NVDA Log Viewer not accessible"))
				return
			self._refreshBookmarks(textCtrl, onComplete=lambda: self._advanceBookmark(textCtrl, forward=forward))
			return

		extCtrl = self._getExternalLogTextControl()
		if not (extCtrl and self.current_log_file):
			gesture.send()
			return
		# Only the plain window handle crosses into the worker; the NVDAObject stays on the main thread.
		windowHandle = extCtrl.windowHandle
		externalLogPath = self.current_log_file
		direction = "next" if forward else "prev"

		def loadBookmarksAsync():
			bookmarks = loadBookmarksFromFile(externalLogPath)
			wx.CallAfter(self._process_external_bookmark_navigation_by_handle, windowHandle, bookmarks, direction)

		threading.Thread(target=loadBookmarksAsync, daemon=True).start()

	@script(description=_("Move to next bookmark in log"), gesture="kb:f2", category=_("LogViewer"))
	def script_moveToNextBookmark(self, gesture):
		self._navigateBookmark(gesture, True)

	@script(description=_("Move to previous bookmark in log"), gesture="kb:shift+f2", category=_("LogViewer"))
	def script_moveToPreviousBookmark(self, gesture):
		self._navigateBookmark(gesture, False)

	def _advanceBookmark(self, textCtrl, forward):
		if not self.bookmarks:
			message(_("No bookmarks found"))
			self.currentBookmark = -1
			return
		wrap = PluginSettings.get().searchWrap
		if forward:
			self.currentBookmark += 1
			if self.currentBookmark >= len(self.bookmarks):
				if wrap:
					self.currentBookmark = 0
					message(_("Wrapping to first bookmark"))
				else:
					self.currentBookmark = len(self.bookmarks) - 1
					message(_("Reached end of bookmarks"))
					return
		else:
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

	def _process_external_bookmark_navigation_by_handle(self, windowHandle, bookmarks, direction):
		# Main thread: the NVDAObject is rebuilt here from the integer handle.
		if self._isTerminated:
			return
		try:
			textCtrl = NVDAWindowObject(windowHandle=windowHandle)
		except (OSError, RuntimeError) as e:
			log.error(f"Error reconstructing external log control from handle: {e}")
			message(_("NVDA Log Viewer not accessible"))
			return
		if not textCtrl:
			message(_("NVDA Log Viewer not accessible"))
			return
		self._process_external_bookmark_navigation(textCtrl, bookmarks, direction)

	def _process_external_bookmark_navigation(self, textCtrl, bookmarks, direction):
		if not bookmarks:
			message(_("No bookmarks found in file"))
			return
		wrap = PluginSettings.get().searchWrap
		caretPos = self.getCaretPosition(textCtrl)
		current_idx = -1
		for i, (start, end, num) in enumerate(bookmarks):
			if start <= caretPos < end:
				current_idx = i
				break
		if direction == "next":
			target_idx = current_idx + 1
			if target_idx >= len(bookmarks):
				if wrap:
					target_idx = 0
					message(_("Wrapping to first bookmark"))
				else:
					message(_("Reached end of bookmarks"))
					return
		else:
			target_idx = current_idx - 1
			if target_idx < 0:
				if wrap:
					target_idx = len(bookmarks) - 1
					message(_("Wrapping to last bookmark"))
				else:
					message(_("Already at first bookmark"))
					return
		self._moveToBookmarkExternal(textCtrl, bookmarks, target_idx)

	def _moveToBookmark(self, textCtrl):
		if not self.bookmarks or self.currentBookmark < 0 or self.currentBookmark >= len(self.bookmarks):
			message(_("No bookmarks available"))
			return
		start_pos, end_pos, bookmark_num = self.bookmarks[self.currentBookmark]

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
			except (RuntimeError, NotImplementedError) as e:
				log.error(f"Error moving to bookmark: {e}")
				message(_("Error moving to bookmark"))

		wx.CallAfter(_move)

	def _moveToBookmarkExternal(self, textCtrl, bookmarks, index):
		if not bookmarks or index < 0 or index >= len(bookmarks):
			message(_("No bookmarks available"))
			return
		start_pos, end_pos, bookmark_num = bookmarks[index]

		def _move():
			try:
				textInfo = textCtrl.makeTextInfo(textInfos.POSITION_ALL)
				textInfo.collapse()
				textInfo.move(textInfos.UNIT_CHARACTER, start_pos)
				textInfo.collapse()
				textInfo.updateSelection()
				message(_("Bookmark {number}").format(number=bookmark_num))
			except (RuntimeError, NotImplementedError) as e:
				log.error(f"Error moving to bookmark in external editor: {e}")
				message(_("Error moving to bookmark"))

		wx.CallAfter(_move)

	def _performFind(self, textCtrl, forward):
		settings = PluginSettings.get()
		manager = self.search_manager
		quickTerm = settings.quickSearchTerms[0] if settings.quickSearchTerms else "error"
		term = manager.lastSearchTerm or quickTerm
		searchType = SearchType.getByName(settings.searchType)
		wrap = settings.searchWrap

		def onReady(status):
			if self._isTerminated:
				return
			if status != "ok":
				message(describeSearchFailure(status))
				return
			self._stepToMatch(textCtrl, forward, wrap)

		manager.requestMatches(textCtrl, term, settings.searchCaseSensitive, searchType, onReady)

	def _stepToMatch(self, textCtrl, forward, wrap, inclusive=False):
		manager = self.search_manager
		if not manager.lastMatches:
			message(_("No matches found for '{term}'").format(term=manager.lastSearchTerm))
			return
		caretPos = self.getCaretPosition(textCtrl)
		if forward:
			index = manager.findNextMatch(caretPos, wrap, inclusive)
		else:
			index = manager.findPrevMatch(caretPos, wrap)
		if index == -1:
			message(_("Reached end of matches") if forward else _("Reached start of matches"))
			return
		matchStart = manager.lastMatches[index][0]
		if forward and (matchStart < caretPos or (matchStart == caretPos and not inclusive)):
			message(_("Wrapping to first match"))
		elif not forward and matchStart >= caretPos:
			message(_("Wrapping to last match"))
		manager.currentMatchIndex = index
		self._moveToQuickSearchResult(textCtrl)

	def _copyErrorBlockAtCurrentMatch(self, textCtrl):
		manager = self.search_manager
		if manager.lastMatches and 0 <= manager.currentMatchIndex < len(manager.lastMatches):
			# Offsets belong to the text the search ran on, so cut the block from that same text.
			allText = manager.lastText
			position = manager.lastMatches[manager.currentMatchIndex][0]
		else:
			allText = get_full_text(textCtrl)
			position = self.getCaretPosition(textCtrl)
		if not allText:
			log.error("Could not retrieve log text")
			message(_("Could not read the log text"))
			return

		startOffset, endOffset, blockType = get_block_at_position(allText, position)
		if startOffset is None:
			message(_("No error block found here"))
			return
		if not self._copyTextToClipboard(allText[startOffset:endOffset]):
			message(_("Could not access clipboard"))
			return
		self._scheduleConfirmationBeep(0)

	@script(description=_("Find next occurrence (single tap) or copy error block (double tap)"), gesture="kb:f3", category=_("LogViewer"))
	def script_findNext(self, gesture):
		if not self.isNVDAViewer():
			gesture.send()
			return
		textCtrl = self.getLogTextControl()
		if not textCtrl:
			message(_("NVDA Log Viewer not accessible"))
			return

		current_time = time.time()
		if current_time - self._findNext_tap_time > self._tap_threshold:
			self._findNext_tap_count = 0
		self._findNext_tap_count += 1
		self._findNext_tap_time = current_time

		if self._findNext_tap_timer:
			self._findNext_tap_timer.Stop()
			self._findNext_tap_timer = None

		def execute_action():
			try:
				if self._findNext_tap_count == 1:
					self._performFind(textCtrl, True)
				elif self._findNext_tap_count >= 2:
					self._copyErrorBlockAtCurrentMatch(textCtrl)
				self._findNext_tap_count = 0
				self._findNext_tap_timer = None
			except RuntimeError as e:
				log.error(f"Error in execute_action: {e}")

		self._findNext_tap_timer = wx.CallLater(int(self._tap_threshold * 1000), execute_action)

	@script(description=_("Find previous occurrence"), gesture="kb:shift+f3", category=_("LogViewer"))
	def script_findPrevious(self, gesture):
		if not self.isNVDAViewer():
			gesture.send()
			return
		textCtrl = self.getLogTextControl()
		if not textCtrl:
			message(_("NVDA Log Viewer not accessible"))
			return
		self._performFind(textCtrl, False)

	def _moveToQuickSearchResult(self, textCtrl):
		announce_total = self.search_manager.newSearchPerformed
		self.search_manager.moveToResult(textCtrl, self.search_manager.currentMatchIndex, announce_total)
		self.search_manager.newSearchPerformed = False

	@script(description=_("Open NVDA log file (prefers nvda-old.log, falls back to nvda.log)"), gesture="kb:NVDA+control+l", category=_("LogViewer"))
	def script_openOldLog(self, gesture):
		def open_log_file():
			try:
				fileToOpen = pickLogPath(True, getCurrentLogPath(), getOldLogPath(), os.path.isfile)
				if not fileToOpen:
					wx.CallAfter(message, _("No NVDA log file found"))
					return
				fileType = _("old log file") if fileToOpen == getOldLogPath() else _("current log file")
				self.current_log_file = fileToOpen
				os.startfile(fileToOpen)
				wx.CallAfter(message, _("Opening {file_type}").format(file_type=fileType))
			except OSError as e:
				log.error(f"Error opening log file: {e}")
				wx.CallAfter(message, _("Failed to open log file"))

		threading.Thread(target=open_log_file, daemon=True).start()

	@script(
		description=_("Copy log since the latest bookmark (single tap), step back through older bookmarks copying each one's segment (double tap), or switch between the current and old log (triple tap)"),
		gesture="kb:control+shift+f2",
		category=_("LogViewer")
	)
	def script_copyFromBookmark(self, gesture):
		current_time = time.time()
		if current_time - self._copyBookmark_tap_time > self._tap_threshold:
			self._copyBookmark_tap_count = 0
		self._copyBookmark_tap_count += 1
		self._copyBookmark_tap_time = current_time

		if self._copyBookmark_tap_timer:
			self._copyBookmark_tap_timer.Stop()
			self._copyBookmark_tap_timer = None

		def execute_action():
			try:
				if self._copyBookmark_tap_count == 1:
					self._startBookmarkCopy(stepBack=False)
				elif self._copyBookmark_tap_count == 2:
					self._startBookmarkCopy(stepBack=True)
				elif self._copyBookmark_tap_count >= 3:
					self._toggleClipboardLogMode()
				self._copyBookmark_tap_count = 0
				self._copyBookmark_tap_timer = None
			except RuntimeError as e:
				log.error(f"Error in copy-from-bookmark action: {e}")

		self._copyBookmark_tap_timer = wx.CallLater(int(self._tap_threshold * 1000), execute_action)

	def _toggleClipboardLogMode(self):
		self._clipboardLogMode = "old" if self._clipboardLogMode == "current" else "current"
		# Bookmark indexes of one file mean nothing in the other.
		self._clipboardBookmarkCursor = None
		logPath = pickLogPath(
			self._clipboardLogMode == "old", getCurrentLogPath(), getOldLogPath(), os.path.isfile
		)
		if not logPath:
			message(_("Log file not found"))
			return
		message(_("Old log") if self._clipboardLogMode == "old" else _("Current log"))

	def _startBookmarkCopy(self, stepBack):
		preferOld = self._clipboardLogMode == "old"
		cursor = self._clipboardBookmarkCursor

		def worker():
			result = buildBookmarkCopy(preferOld, cursor, stepBack)
			wx.CallAfter(self._finishBookmarkCopy, result)

		threading.Thread(target=worker, daemon=True).start()

	def _finishBookmarkCopy(self, result):
		# Main thread: clipboard and speech must not be touched from the reading worker.
		if self._isTerminated:
			return
		if result.status == "noLog":
			message(_("No NVDA log file found"))
			return
		if result.status == "readFailed":
			message(_("Failed to read log file"))
			return
		self._clipboardBookmarkCursor = result.cursorIndex
		lineCount = countLines(result.text)
		if result.status == "wholeLog":
			log.info(f"Whole log copy: file={result.logPath}, characters={len(result.text)}, lines={lineCount}")
			if not result.text or not self._copyTextToClipboard(result.text):
				message(_("Could not access clipboard"))
				return
			message(_("No bookmark yet, whole log copied, {count} lines").format(count=lineCount))
			self._scheduleConfirmationBeep(200)
			return
		log.info(
			f"Bookmark copy: file={result.logPath}, bookmark={result.bookmarkNumber}, "
			f"characters={len(result.text)}, lines={lineCount}"
		)
		if result.wrapped:
			message(_("Wrapping to latest bookmark"))
		if not result.text:
			message(_("Bookmark {number} has no log lines").format(number=result.bookmarkNumber))
			return
		if not self._copyTextToClipboard(result.text):
			message(_("Could not access clipboard"))
			return
		message(_("Bookmark {number}, {count} lines").format(number=result.bookmarkNumber, count=lineCount))
		self._scheduleConfirmationBeep(200)
