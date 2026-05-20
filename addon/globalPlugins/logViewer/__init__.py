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

from .config_manager import initConfiguration, SearchHistory
from .search_logic import SearchType, SearchManager, LogSearchDialog, fIsLogViewer, get_block_at_position

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
		self._logViewerWeakRef = None
		self.bookmarks = []
		self.currentBookmark = -1
		self.searchDialog = None
		self.bookmarkLock = threading.Lock()
		self.lastBookmarkRefreshTime = 0
		self.current_log_file = None
		threading.Thread(target=self._addCrashBookmarkIfNeeded, daemon=True).start()
		self.bookmarkCount = 1
		config.conf["LogViewerPlugin"]["bookmarkCount"] = 1
		config.conf.save()

		self.search_manager = SearchManager()
		self.search_manager.lastSearchTerm = "error"

		self.lastSearchTerm = ""
		self.lastMatches = []
		self.currentMatchIndex = -1
		self._lastSearchCaseSensitive = None
		self._lastSearchType = None

		self._findNext_tap_time = 0
		self._findNext_tap_count = 0
		self._tap_threshold = 0.5
		self._findNext_tap_timer = None

	def terminate(self):
		if hasattr(self, '_findNext_tap_timer') and self._findNext_tap_timer:
			self._findNext_tap_timer.Stop()
			self._findNext_tap_timer = None
		try:
			if hasattr(self, 'searchDialog') and self.searchDialog:
				if self.searchDialog.dialogOpen:
					self.searchDialog.Destroy()
				self.searchDialog = None
			self.bookmarks = None
			self._logViewerWeakRef = None
		except Exception:
			pass

	def _addCrashBookmarkIfNeeded(self):
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
			bookmark_line = f"\n{self.bookmarkString.format(self.bookmarkCount)}\n"
			with open(old_log, 'a', encoding='utf-8') as f:
				f.write(bookmark_line)
			log.info(f"Added crash bookmark to old log: {bookmark_line.strip()}")
			self.bookmarkCount += 1
			config.conf["LogViewerPlugin"]["bookmarkCount"] = self.bookmarkCount
			config.conf.save()
		except Exception as e:
			log.error(f"Error adding crash bookmark: {e}")

	def isNVDAViewer(self):
		try:
			focusObj = api.getFocusObject()
			if not focusObj:
				return False
			return self.isNVDAViewerObject(focusObj)
		except Exception as e:
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
		except Exception as e:
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
		except Exception:
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
		except Exception as e:
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
			except Exception:
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

				self.searchDialog = LogSearchDialog(topWin, textCtrl, self)
				gui.mainFrame.prePopup()
				self.searchDialog.Show()

				def _post_popup():
					gui.mainFrame.postPopup()
				core.callLater(100, _post_popup)
			except Exception as e:
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

		if self.isNVDAViewer():
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
			return

		extCtrl = self._getExternalLogTextControl()
		if extCtrl and self.current_log_file:
			def load_bookmarks_async():
				bookmarks = self._refreshBookmarksFromFile(self.current_log_file)
				wx.CallAfter(self._process_external_bookmark_navigation, extCtrl, bookmarks, "next")
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
			return

		extCtrl = self._getExternalLogTextControl()
		if extCtrl and self.current_log_file:
			def load_bookmarks_async():
				bookmarks = self._refreshBookmarksFromFile(self.current_log_file)
				wx.CallAfter(self._process_external_bookmark_navigation, extCtrl, bookmarks, "prev")
			threading.Thread(target=load_bookmarks_async, daemon=True).start()
		else:
			gesture.send()

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
				except Exception as e:
					log.error(f"Error moving to bookmark: {e}")
					message(_("Error moving to bookmark"))
			wx.CallAfter(_move)
		except Exception as e:
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
				except Exception as e:
					log.error(f"Error moving to bookmark in external editor: {e}")
					message(_("Error moving to bookmark"))
			wx.CallAfter(_move)
		except Exception as e:
			log.error(f"Error in _moveToBookmarkExternal: {e}")

	def _performFindNext(self, textCtrl):
		if not self.search_manager.lastSearchTerm:
			self.search_manager.lastSearchTerm = "error"
			caseSensitive = config.conf["LogViewerPlugin"]["searchCaseSensitivity"]
			searchType = SearchType.getByName(config.conf["LogViewerPlugin"]["searchType"])
			if not self.search_manager.doQuickSearch(textCtrl, "error", caseSensitive, searchType):
				wx.CallAfter(message, _("No matches found for 'error'"))
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
			self.search_manager.lastSearchTerm = "error"
			caseSensitive = config.conf["LogViewerPlugin"]["searchCaseSensitivity"]
			searchType = SearchType.getByName(config.conf["LogViewerPlugin"]["searchType"])
			if not self.search_manager.doQuickSearch(textCtrl, "error", caseSensitive, searchType):
				wx.CallAfter(message, _("No matches found for 'error'"))
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
				start_pos, _ = self.search_manager.lastMatches[self.search_manager.currentMatchIndex]
				pos = start_pos
			else:
				pos = self.getCaretPosition(textCtrl)

			try:
				textInfo = textCtrl.makeTextInfo(textInfos.POSITION_ALL)
				all_text = textInfo.text
			except Exception as e:
				log.error(f"Error getting full text: {e}")
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
					except Exception as e:
						log.error(f"Error during beep: {e}")
				core.callLater(0, play_beep)
			else:
				log.error("Could not open clipboard")
		except Exception as e:
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
			except Exception as e:
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
					wx.CallAfter(message, _("No NVDA log file found"))
					return
				self.current_log_file = file_to_open
				if sys.platform.startswith("win"):
					os.startfile(file_to_open)
				else:
					subprocess.run(["xdg-open", file_to_open], check=True)
				wx.CallAfter(message, _("Opening {file_type}").format(file_type=message_type))
			except Exception as e:
				log.error(f"Error opening log file: {e}")
				wx.CallAfter(message, _("Failed to open log file"))

		threading.Thread(target=open_log_file, daemon=True).start()