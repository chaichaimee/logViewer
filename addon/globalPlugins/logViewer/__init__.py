# __init__.py
# Copyright (C) 2026 Chai Chaimee
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
from NVDAObjects.window import Window as NVDAWindowObject
from logHandler import log
import addonHandler
import config
import core
import time
import threading
import gui.logViewer
import os
import subprocess
import tempfile
import ctypes
from ctypes import wintypes
import tones
import weakref
import sys

from .config_manager import initConfiguration, SearchHistory, PluginSettings
from .search_logic import SearchType, SearchManager, LogSearchDialog, get_block_at_position, get_full_text, fIsLogViewer
from .settings_panel import LogViewerSettingsPanel

addonHandler.initTranslation()

GA_PARENT = 1
user32 = ctypes.windll.user32
user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
user32.GetAncestor.restype = wintypes.HWND


class GlobalPlugin(GlobalPlugin):
	bookmarkString = "BOOKMARK {0}"

	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
		initConfiguration()
		SearchHistory.get()
		self.bookmarkCount = config.conf["LogViewerPlugin"].get("bookmarkCount", 1)
		self._resetBookmarkCountOnNewBoot()
		self._logViewerWeakRef = None
		self.bookmarks = []
		self.currentBookmark = -1
		self.searchDialog = None
		self.bookmarkLock = threading.Lock()
		self.lastBookmarkRefreshTime = 0
		self.current_log_file = None
		threading.Thread(target=self._addCrashBookmarkIfNeeded, daemon=True).start()

		self.search_manager = SearchManager()
		# lastSearchTerm intentionally starts empty so the first F3 press picks
		# up the configured quick-search default term (see
		# _getDefaultQuickSearchTerm) instead of a hardcoded "error" that
		# bypassed the settings panel's quick-search term list.

		self.lastSearchTerm = ""
		self.lastMatches = []
		self.currentMatchIndex = -1
		self._lastSearchCaseSensitive = None
		self._lastSearchType = None

		self._findNext_tap_time = 0
		self._findNext_tap_count = 0
		self._tap_threshold = 0.5
		self._findNext_tap_timer = None

		# State for the ctrl+shift+f2 "copy from bookmark" feature: which
		# bookmark, by index into the sorted bookmark list, the double-tap
		# gesture last stepped back to. None means "not yet used this
		# session" -- the first double-tap after that steps back one from
		# whatever the latest bookmark is.
		self._copyBookmark_tap_time = 0
		self._copyBookmark_tap_count = 0
		self._copyBookmark_tap_timer = None
		self._clipboardBookmarkCursor = None
		# Which log file the ctrl+shift+f2 feature reads from. Defaults to the
		# current session's nvda.log, since that's where the user is actively
		# inserting new bookmarks; nvda-old.log (the previous session) is only
		# used when explicitly switched to via triple-tap.
		self._clipboardLogMode = "current"

		if LogViewerSettingsPanel not in gui.settingsDialogs.NVDASettingsDialog.categoryClasses:
			gui.settingsDialogs.NVDASettingsDialog.categoryClasses.append(LogViewerSettingsPanel)

	def _getSystemBootTimestamp(self):
		"""Returns an approximate Unix timestamp for when Windows last booted,
		derived from GetTickCount64 (milliseconds elapsed since boot). This
		value is stable across NVDA restarts but changes on every full system
		boot, which lets bookmark numbering be reset only on an actual reboot
		rather than every time NVDA itself restarts. Requires no elevated
		privileges, unlike WMI-based boot time queries."""
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
		# A few seconds of slack absorbs clock rounding between GetTickCount64
		# reads taken in different NVDA sessions of the same boot; anything
		# larger means the system was actually restarted since we last saw it.
		if abs(currentBootTimestamp - float(storedBootTimestamp)) > 5.0:
			self.bookmarkCount = 1
			config.conf["LogViewerPlugin"]["bookmarkCount"] = 1
			config.conf["LogViewerPlugin"]["lastBootTimestamp"] = currentBootTimestamp
			config.conf.save()

	def terminate(self):
		if hasattr(self, '_findNext_tap_timer') and self._findNext_tap_timer:
			self._findNext_tap_timer.Stop()
			self._findNext_tap_timer = None
		if hasattr(self, '_copyBookmark_tap_timer') and self._copyBookmark_tap_timer:
			self._copyBookmark_tap_timer.Stop()
			self._copyBookmark_tap_timer = None
		try:
			if hasattr(self, 'searchDialog') and self.searchDialog:
				if self.searchDialog.dialogOpen:
					self.searchDialog.Destroy()
				self.searchDialog = None
			self.bookmarks = None
			self._logViewerWeakRef = None
			if LogViewerSettingsPanel in gui.settingsDialogs.NVDASettingsDialog.categoryClasses:
				gui.settingsDialogs.NVDASettingsDialog.categoryClasses.remove(LogViewerSettingsPanel)
		except (AttributeError, ValueError, RuntimeError) as e:
			log.error(f"Error during terminate: {e}")

	def _addCrashBookmarkIfNeeded(self):
		# Runs on a background thread. Disk reads and the append write are safe
		# here; the shared bookmarkCount counter and config.conf mutation/save
		# are NOT, since config.conf is bound to the main thread. That part is
		# marshaled back via core.callLater in _commitCrashBookmarkCount.
		try:
			temp_dir = tempfile.gettempdir()
			old_log = os.path.join(temp_dir, "nvda-old.log")
			if not os.path.exists(old_log):
				return
			with open(old_log, 'r', encoding='utf-8', errors='ignore') as f:
				lines = f.readlines()
			if not lines:
				return
			last_line = lines[-1].strip()
			crash_indicators = ["Traceback", "ERROR - unhandled exception", "CRASH"]
			is_crash = any(indicator in last_line for indicator in crash_indicators)
			if not is_crash:
				return
			has_bookmark = any(line.strip().startswith("BOOKMARK") for line in lines[-5:])
			if has_bookmark:
				return
			pending_bookmark_number = self.bookmarkCount
			bookmark_line = f"\n{self.bookmarkString.format(pending_bookmark_number)}\n"
			with open(old_log, 'a', encoding='utf-8') as f:
				f.write(bookmark_line)
			log.info(f"Added crash bookmark to old log: {bookmark_line.strip()}")
			core.callLater(0, self._commitCrashBookmarkCount, pending_bookmark_number + 1)
		except OSError as e:
			log.error(f"Error adding crash bookmark: {e}")

	def _commitCrashBookmarkCount(self, newCount):
		# Runs on the main thread via core.callLater, so mutating config.conf
		# and saving it here is safe.
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

	def _refreshBookmarksFromFile(self, file_path):
		bookmarks = []
		try:
			with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
				content = f.read()
			pattern = re.compile(r"BOOKMARK (\d+)")
			for match in pattern.finditer(content):
				start, end = match.span()
				num = int(match.group(1))
				bookmarks.append((start, end, num))
			bookmarks.sort(key=lambda x: x[0])
		except OSError as e:
			log.error(f"Error reading bookmarks from file {file_path}: {e}")
		return bookmarks

	@script(description=_("Search in NVDA Log Viewer"), gesture="kb:control+f", category=_("LogViewer"))
	def script_searchInLogViewer(self, gesture):
		if not self.isNVDAViewer():
			gesture.send()
			return
		textCtrl = self.getLogTextControl()
		if not textCtrl:
			message(_("NVDA Log Viewer not accessible"))
			return

		if self.searchDialog is not None:
			try:
				if self.searchDialog.dialogOpen:
					self.searchDialog.Raise()
					self.searchDialog.searchBox.SetFocus()
					return
				else:
					self.searchDialog = None
			except RuntimeError:
				self.searchDialog = None

		def showDialog():
			try:
				if not wx.IsMainThread():
					wx.CallAfter(showDialog)
					return

				app = wx.GetApp()
				if not app:
					log.error("No wx.App instance available")
					return
				topWin = app.GetTopWindow()
				if not topWin:
					log.error("No top window available")
					return

				dlg = LogSearchDialog(topWin, textCtrl, self)
				self.searchDialog = dlg
				gui.mainFrame.prePopup()

				# Call postPopup only once the dialog has genuinely finished
				# closing (bound to the real destroy event), instead of guessing
				# at a fixed millisecond delay that could fire while the dialog
				# is still open or before the OS has restored focus.
				def _onDialogDestroy(evt):
					evt.Skip()
					if evt.GetEventObject() is dlg:
						core.callLater(0, gui.mainFrame.postPopup)
				dlg.Bind(wx.EVT_WINDOW_DESTROY, _onDialogDestroy)

				dlg.Show()
			except RuntimeError as e:
				log.error(f"Error opening search dialog: {str(e)}")
				message(_("Failed to open search dialog"))

		wx.CallAfter(showDialog)

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
		# The heavy work (full-text retrieval plus regex scanning of the whole
		# log) is offloaded to a background thread so it cannot trip NVDA's
		# main-thread watchdog on large log files. Only the resulting bookmark
		# list is marshaled back to the main thread.
		current_time = time.time()
		if current_time - self.lastBookmarkRefreshTime < 0.1 and self.bookmarks:
			if onComplete:
				onComplete()
			return

		def worker():
			foundBookmarks = []
			try:
				all_log_text = get_full_text(textCtrl)
				if not all_log_text or not all_log_text.strip():
					wx.CallAfter(message, _("Log is empty"))
					wx.CallAfter(self._applyRefreshedBookmarks, [], onComplete)
					return
				bookmark_pattern = re.compile(r"BOOKMARK (\d+)")
				for match in bookmark_pattern.finditer(all_log_text):
					start_pos, end_pos = match.span()
					bookmark_num = int(match.group(1))
					foundBookmarks.append((start_pos, end_pos, bookmark_num))
				foundBookmarks.sort(key=lambda x: x[0])
			except re.error as e:
				log.error(f"Error refreshing bookmarks: {e}")
				foundBookmarks = []
			wx.CallAfter(self._applyRefreshedBookmarks, foundBookmarks, onComplete)

		threading.Thread(target=worker, daemon=True).start()

	def _applyRefreshedBookmarks(self, foundBookmarks, onComplete):
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

	@script(description=_("Move to next bookmark in log"), gesture="kb:f2", category=_("LogViewer"))
	def script_moveToNextBookmark(self, gesture):
		if self.isInBookmarkConflictingApp():
			gesture.send()
			return

		if self.isNVDAViewer():
			textCtrl = self.getLogTextControl()
			if not textCtrl:
				message(_("NVDA Log Viewer not accessible"))
				return
			self._refreshBookmarks(textCtrl, onComplete=lambda: self._advanceBookmark(textCtrl, forward=True))
			return

		extCtrl = self._getExternalLogTextControl()
		if extCtrl and self.current_log_file:
			# Extract only the primitive window handle on the main thread before
			# handing work to the background thread. The NVDAObject itself (a COM
			# wrapper created on NVDA's Single-Threaded Apartment) must never be
			# captured by a background thread's closure.
			windowHandle = extCtrl.windowHandle
			externalLogPath = self.current_log_file

			def load_bookmarks_async():
				bookmarks = self._refreshBookmarksFromFile(externalLogPath)
				wx.CallAfter(self._process_external_bookmark_navigation_by_handle, windowHandle, bookmarks, "next")
			threading.Thread(target=load_bookmarks_async, daemon=True).start()
		else:
			gesture.send()

	@script(description=_("Move to previous bookmark in log"), gesture="kb:shift+f2", category=_("LogViewer"))
	def script_moveToPreviousBookmark(self, gesture):
		if self.isInBookmarkConflictingApp():
			gesture.send()
			return

		if self.isNVDAViewer():
			textCtrl = self.getLogTextControl()
			if not textCtrl:
				message(_("NVDA Log Viewer not accessible"))
				return
			self._refreshBookmarks(textCtrl, onComplete=lambda: self._advanceBookmark(textCtrl, forward=False))
			return

		extCtrl = self._getExternalLogTextControl()
		if extCtrl and self.current_log_file:
			windowHandle = extCtrl.windowHandle
			externalLogPath = self.current_log_file

			def load_bookmarks_async():
				bookmarks = self._refreshBookmarksFromFile(externalLogPath)
				wx.CallAfter(self._process_external_bookmark_navigation_by_handle, windowHandle, bookmarks, "prev")
			threading.Thread(target=load_bookmarks_async, daemon=True).start()
		else:
			gesture.send()

	def _advanceBookmark(self, textCtrl, forward):
		if not self.bookmarks:
			message(_("No bookmarks found"))
			self.currentBookmark = -1
			return
		wrap = config.conf["LogViewerPlugin"]["searchWrap"]
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
		# Runs on the main thread (via wx.CallAfter). Reconstruct the NVDAObject
		# here, from the plain integer handle, rather than ever having passed the
		# COM-backed object itself through the background thread.
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
		wrap = config.conf["LogViewerPlugin"]["searchWrap"]
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
				except (RuntimeError, NotImplementedError) as e:
					log.error(f"Error moving to bookmark: {e}")
					message(_("Error moving to bookmark"))
			wx.CallAfter(_move)
		except RuntimeError as e:
			log.error(f"Error in _moveToBookmark: {e}")

	def _moveToBookmarkExternal(self, textCtrl, bookmarks, index):
		if not bookmarks or index < 0 or index >= len(bookmarks):
			message(_("No bookmarks available"))
			return
		start_pos, end_pos, bookmark_num = bookmarks[index]
		try:
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
		except RuntimeError as e:
			log.error(f"Error in _moveToBookmarkExternal: {e}")

	def _getDefaultQuickSearchTerm(self):
		terms = PluginSettings.get().quickSearchTerms
		return terms[0] if terms else "error"

	def _performFindNext(self, textCtrl):
		if not self.search_manager.lastSearchTerm:
			defaultTerm = self._getDefaultQuickSearchTerm()
			self.search_manager.lastSearchTerm = defaultTerm
			caseSensitive = config.conf["LogViewerPlugin"]["searchCaseSensitivity"]
			searchType = SearchType.getByName(config.conf["LogViewerPlugin"]["searchType"])
			if not self.search_manager.doQuickSearch(textCtrl, defaultTerm, caseSensitive, searchType):
				wx.CallAfter(message, _("No matches found for '{term}'").format(term=defaultTerm))
				return

		caseSensitive = config.conf["LogViewerPlugin"]["searchCaseSensitivity"]
		searchType = SearchType.getByName(config.conf["LogViewerPlugin"]["searchType"])
		wrap = config.conf["LogViewerPlugin"]["searchWrap"]

		if (not self.search_manager.lastMatches or
				caseSensitive != self.search_manager.lastCaseSensitive or
				searchType != self.search_manager.lastSearchType):
			if not self.search_manager.doQuickSearch(textCtrl, self.search_manager.lastSearchTerm, caseSensitive, searchType):
				wx.CallAfter(message, _("No matches found"))
				return

		if not self.search_manager.lastMatches:
			wx.CallAfter(message, _("No matches found"))
			return

		caretPos = self.getCaretPosition(textCtrl)
		idx = self.search_manager.findNextMatch(caretPos, wrap)
		if idx == -1:
			wx.CallAfter(message, _("No matches found"))
			return

		self.search_manager.currentMatchIndex = idx
		self._moveToQuickSearchResult(textCtrl)

	def _performFindPrevious(self, textCtrl):
		if not self.search_manager.lastSearchTerm:
			defaultTerm = self._getDefaultQuickSearchTerm()
			self.search_manager.lastSearchTerm = defaultTerm
			caseSensitive = config.conf["LogViewerPlugin"]["searchCaseSensitivity"]
			searchType = SearchType.getByName(config.conf["LogViewerPlugin"]["searchType"])
			if not self.search_manager.doQuickSearch(textCtrl, defaultTerm, caseSensitive, searchType):
				wx.CallAfter(message, _("No matches found for '{term}'").format(term=defaultTerm))
				return

		caseSensitive = config.conf["LogViewerPlugin"]["searchCaseSensitivity"]
		searchType = SearchType.getByName(config.conf["LogViewerPlugin"]["searchType"])
		wrap = config.conf["LogViewerPlugin"]["searchWrap"]

		if (not self.search_manager.lastMatches or
				caseSensitive != self.search_manager.lastCaseSensitive or
				searchType != self.search_manager.lastSearchType):
			if not self.search_manager.doQuickSearch(textCtrl, self.search_manager.lastSearchTerm, caseSensitive, searchType):
				wx.CallAfter(message, _("No matches found"))
				return

		if not self.search_manager.lastMatches:
			wx.CallAfter(message, _("No matches found"))
			return

		caretPos = self.getCaretPosition(textCtrl)
		idx = self.search_manager.findPrevMatch(caretPos, wrap)
		if idx == -1:
			wx.CallAfter(message, _("No matches found"))
			return

		self.search_manager.currentMatchIndex = idx
		self._moveToQuickSearchResult(textCtrl)

	def _copyErrorBlockAtCurrentMatch(self, textCtrl):
		try:
			if self.search_manager.lastMatches and self.search_manager.currentMatchIndex >= 0:
				start_pos, _unused_end = self.search_manager.lastMatches[self.search_manager.currentMatchIndex]
				pos = start_pos
			else:
				pos = self.getCaretPosition(textCtrl)

			all_text = get_full_text(textCtrl)
			if not all_text:
				log.error("Could not retrieve log text")
				return

			start_offset, end_offset, block_type = get_block_at_position(all_text, pos)
			if start_offset is None:
				log.info("No block found at current position")
				return

			block_text = all_text[start_offset:end_offset]

			if wx.TheClipboard.Open():
				wx.TheClipboard.SetData(wx.TextDataObject(block_text))
				wx.TheClipboard.Close()

				def play_beep():
					try:
						tones.beep(440, 100)
						log.info("Block copied, beep played")
					except RuntimeError as e:
						log.error(f"Error during beep: {e}")
				core.callLater(0, play_beep)
			else:
				log.error("Could not open clipboard")
		except (RuntimeError, NotImplementedError) as e:
			log.error(f"Unexpected error in _copyErrorBlockAtCurrentMatch: {e}")

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
					self._performFindNext(textCtrl)
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
		self._performFindPrevious(textCtrl)

	def _moveToQuickSearchResult(self, textCtrl):
		announce_total = self.search_manager.newSearchPerformed
		self.search_manager.moveToResult(textCtrl, self.search_manager.currentMatchIndex, announce_total)
		self.search_manager.newSearchPerformed = False

	@script(description=_("Open NVDA log file (prefers nvda-old.log, falls back to nvda.log)"), gesture="kb:NVDA+control+l", category=_("LogViewer"))
	def script_openOldLog(self, gesture):
		def open_log_file():
			try:
				temp_dir = tempfile.gettempdir()
				old_log_path = os.path.join(temp_dir, "nvda-old.log")
				current_log_path = os.path.join(temp_dir, "nvda.log")

				if os.path.exists(old_log_path):
					file_to_open = old_log_path
					message_type = _("old log file")
				elif os.path.exists(current_log_path):
					file_to_open = current_log_path
					message_type = _("current log file")
				else:
					core.callLater(0, message, _("No NVDA log file found"))
					return

				self.current_log_file = file_to_open

				if sys.platform.startswith("win"):
					os.startfile(file_to_open)
				else:
					subprocess.run(["xdg-open", file_to_open], check=True)

				core.callLater(0, message, _("Opening {file_type}").format(file_type=message_type))
			except OSError as e:
				log.error(f"Error opening log file: {e}")
				core.callLater(0, message, _("Failed to open log file"))

		threading.Thread(target=open_log_file, daemon=True).start()

	def _resolveLogFilePathForClipboard(self):
		"""Locates the NVDA log file to read bookmarks from directly on disk,
		without requiring the NVDA log viewer to be open. Which file is
		preferred depends on self._clipboardLogMode (toggled via triple-tap):
		"current" prefers nvda.log (where new bookmarks are actively being
		added this session), "old" prefers nvda-old.log (the previous
		session). Either way, falls back to whichever file actually exists if
		the preferred one doesn't."""
		temp_dir = tempfile.gettempdir()
		old_log_path = os.path.join(temp_dir, "nvda-old.log")
		current_log_path = os.path.join(temp_dir, "nvda.log")
		if self._clipboardLogMode == "old":
			if os.path.exists(old_log_path):
				return old_log_path
			if os.path.exists(current_log_path):
				return current_log_path
			return None
		if os.path.exists(current_log_path):
			return current_log_path
		if os.path.exists(old_log_path):
			return old_log_path
		return None

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
					self._copyLatestBookmarkSegment()
				elif self._copyBookmark_tap_count == 2:
					self._stepBackAndCopyBookmarkSegment()
				elif self._copyBookmark_tap_count >= 3:
					self._toggleClipboardLogMode()
				self._copyBookmark_tap_count = 0
				self._copyBookmark_tap_timer = None
			except RuntimeError as e:
				log.error(f"Error in copy-from-bookmark action: {e}")

		self._copyBookmark_tap_timer = wx.CallLater(int(self._tap_threshold * 1000), execute_action)

	def _toggleClipboardLogMode(self):
		self._clipboardLogMode = "old" if self._clipboardLogMode == "current" else "current"
		# The bookmark list differs between the two files, so a cursor
		# position from one file has no valid meaning in the other.
		self._clipboardBookmarkCursor = None
		log_path = self._resolveLogFilePathForClipboard()
		if not log_path:
			message(_("Log file not found"))
			return
		label = _("Old log") if self._clipboardLogMode == "old" else _("Current log")
		message(label)

	def _copyLatestBookmarkSegment(self):
		def worker():
			log_path = self._resolveLogFilePathForClipboard()
			if not log_path:
				core.callLater(0, message, _("No NVDA log file found"))
				return
			bookmarks = self._refreshBookmarksFromFile(log_path)
			if not bookmarks:
				core.callLater(0, message, _("No bookmarks found in log"))
				return
			cursorIndex = len(bookmarks) - 1
			self._clipboardBookmarkCursor = cursorIndex
			wx.CallAfter(self._copyBookmarkSegmentToClipboard, log_path, bookmarks, cursorIndex)
		threading.Thread(target=worker, daemon=True).start()

	def _stepBackAndCopyBookmarkSegment(self):
		def worker():
			log_path = self._resolveLogFilePathForClipboard()
			if not log_path:
				core.callLater(0, message, _("No NVDA log file found"))
				return
			bookmarks = self._refreshBookmarksFromFile(log_path)
			if not bookmarks:
				core.callLater(0, message, _("No bookmarks found in log"))
				return
			if self._clipboardBookmarkCursor is None or self._clipboardBookmarkCursor >= len(bookmarks):
				baseCursor = len(bookmarks) - 1
			else:
				baseCursor = self._clipboardBookmarkCursor
			newCursor = baseCursor - 1
			if newCursor < 0:
				# Loop back to the latest bookmark rather than dead-ending at
				# the first one, matching the wrap-around convention already
				# used by F2/Shift+F2 bookmark navigation elsewhere.
				newCursor = len(bookmarks) - 1
				core.callLater(0, message, _("Wrapping to latest bookmark"))
			self._clipboardBookmarkCursor = newCursor
			wx.CallAfter(self._copyBookmarkSegmentToClipboard, log_path, bookmarks, newCursor)
		threading.Thread(target=worker, daemon=True).start()

	def _copyBookmarkSegmentToClipboard(self, log_path, bookmarks, cursorIndex):
		# Runs on the main thread via wx.CallAfter: the clipboard and speech
		# calls below must not happen on the background thread that read the
		# file and scanned it for bookmark markers.
		try:
			with open(log_path, 'r', encoding='utf-8', errors='ignore') as f:
				content = f.read()
		except OSError as e:
			log.error(f"Error reading log file for bookmark clipboard copy: {e}")
			message(_("Failed to read log file"))
			return

		start_offset, end_offset, bookmark_num = bookmarks[cursorIndex]
		segment_start = end_offset
		if cursorIndex + 1 < len(bookmarks):
			next_start, _next_end, _next_num = bookmarks[cursorIndex + 1]
			segment_end = next_start
		else:
			segment_end = len(content)

		segment_text = content[segment_start:segment_end].strip("\n")

		if not wx.TheClipboard.Open():
			log.error("Could not open clipboard for bookmark segment copy")
			message(_("Could not access clipboard"))
			return
		wx.TheClipboard.SetData(wx.TextDataObject(segment_text))
		wx.TheClipboard.Close()
		message(_("Bookmark {number}").format(number=bookmark_num))

		def play_beep():
			try:
				tones.beep(440, 100)
			except RuntimeError as e:
				log.error(f"Error during beep: {e}")
		core.callLater(200, play_beep)
