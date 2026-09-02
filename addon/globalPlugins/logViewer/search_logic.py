# search_logic.py
# Copyright (C) 2026 Chai Chaimee
# Licensed under GNU General Public License. See COPYING.txt for details.

import wx
import api
import gui
import textInfos
import re
import controlTypes
import core
import ui
from logHandler import log
import addonHandler
from enum import Enum, unique
import threading
import ctypes
from ctypes import wintypes
import config

addonHandler.initTranslation()

GA_PARENT = 1
user32 = ctypes.windll.user32
user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
user32.GetAncestor.restype = wintypes.HWND


def fIsLogViewer(obj):
	if obj is None:
		return False
	if not hasattr(gui, 'logViewer') or gui.logViewer.logViewer is None:
		return False
	if obj.role == controlTypes.Role.PANE:
		hParent = obj.windowHandle
	else:
		hParent = user32.GetAncestor(obj.windowHandle, GA_PARENT)
	try:
		hLogViewer = gui.logViewer.logViewer.GetHandle()
		isLogViewer = hLogViewer == hParent
	except (AttributeError, RuntimeError):
		isLogViewer = False
	return isLogViewer


def get_full_text(textCtrl):
	if fIsLogViewer(textCtrl):
		try:
			# Ensure logViewer and outputCtrl are fully initialized before accessing
			if (hasattr(gui.logViewer, 'logViewer') and
				gui.logViewer.logViewer is not None and
				hasattr(gui.logViewer.logViewer, 'outputCtrl') and
				gui.logViewer.logViewer.outputCtrl is not None):
				text = gui.logViewer.logViewer.outputCtrl.GetValue()
				if text.strip():
					log.debug(f"Got log text via outputCtrl, length: {len(text)}")
					return text
				else:
					log.debug("outputCtrl returned empty text, falling back to makeTextInfo")
			else:
				log.debug("logViewer or outputCtrl not ready, falling back to makeTextInfo")
		except (AttributeError, RuntimeError) as e:
			log.error(f"outputCtrl.GetValue failed: {e}, falling back to makeTextInfo")
		# Fallback to IAccessible textInfo
		try:
			ti = textCtrl.makeTextInfo(textInfos.POSITION_ALL)
			fallback_text = ti.text
			log.debug(f"Fallback text length: {len(fallback_text)}")
			return fallback_text
		except (RuntimeError, NotImplementedError) as e2:
			log.error(f"Fallback makeTextInfo also failed: {e2}")
			return ""
	else:
		try:
			ti = textCtrl.makeTextInfo(textInfos.POSITION_ALL)
			return ti.text
		except (RuntimeError, NotImplementedError) as e:
			log.error(f"Error getting text from text control: {e}")
			return ""


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
		return 0

	@staticmethod
	def getByName(name):
		for type in SearchType:
			if type.name == name:
				return type
		return SearchType.NORMAL

	@staticmethod
	def getSearchTypes():
		return [_(i.value) for i in SearchType]


def extract_line_from_textctrl(textCtrl, position):
	try:
		# Must use the same text source that produced the match offset
		# (get_full_text), not a fresh raw COM call. The NVDA log viewer's
		# COM-based text and its outputCtrl.GetValue() text are not guaranteed
		# to be the same length/content, so mixing sources here makes
		# "position" point at the wrong line -- this was the cause of quick
		# search always reporting line 1 regardless of actual position.
		allText = get_full_text(textCtrl)
		if not allText:
			return 0, ""
		line_num = allText.count('\n', 0, position) + 1
		line_start = allText.rfind('\n', 0, position) + 1
		line_end = allText.find('\n', position)
		if line_end == -1:
			line_end = len(allText)
		line_text = allText[line_start:line_end].strip()
		return line_num, line_text
	except (RuntimeError, NotImplementedError) as e:
		log.error(f"Error extracting line: {e}")
		return 0, ""


def _is_excluded_error_line(line_text):
	excluded_keywords = ["alertForSpellingErrors", "reportSpellingErrors", "Search history initialized"]
	for kw in excluded_keywords:
		if kw in line_text:
			return True
	return False


def get_block_at_position(text, pos):
	line_start = text.rfind('\n', 0, pos) + 1
	line_end = text.find('\n', pos)
	if line_end == -1:
		line_end = len(text)

	scan_start = line_start
	block_type = None
	while scan_start > 0:
		prev_line_start = text.rfind('\n', 0, scan_start - 1) + 1
		prev_line_end = scan_start - 1
		if prev_line_start >= len(text):
			break
		line_content = text[prev_line_start:prev_line_end]
		if "ERROR -" in line_content:
			block_type = "ERROR"
			scan_start = prev_line_start
			break
		if "WARNING -" in line_content:
			block_type = "WARNING"
			scan_start = prev_line_start
			break
		if line_content.strip().startswith("Traceback"):
			block_type = "TRACEBACK"
			scan_start = prev_line_start
			break
		scan_start = prev_line_start
		if scan_start == 0:
			break

	if block_type is None:
		return None, None, None

	block_start = scan_start
	block_end = line_end
	current_pos = line_end
	while current_pos < len(text):
		next_newline = text.find('\n', current_pos + 1)
		if next_newline == -1:
			next_newline = len(text)
		next_line = text[current_pos:next_newline]
		if block_type in ("ERROR", "WARNING"):
			if (next_line.startswith("INFO -") or
				next_line.startswith("WARNING -") or
				next_line.startswith("ERROR -") or
				next_line.startswith("DEBUG -") or
				next_line.strip() == ""):
				break
		else:
			if (next_line.startswith("INFO -") or
				next_line.startswith("WARNING -") or
				next_line.startswith("ERROR -") or
				next_line.startswith("DEBUG -") or
				next_line.strip() == "" or
				(not next_line.startswith(" ") and not next_line.startswith("\t"))):
				break
		block_end = next_newline
		current_pos = next_newline

	return block_start, block_end, block_type


class SearchManager:
	def __init__(self):
		self.lastSearchTerm = ""
		self.lastMatches = []
		self.currentMatchIndex = -1
		self.lastCaseSensitive = None
		self.lastSearchType = None
		self.newSearchPerformed = False

	def doQuickSearch(self, textCtrl, term, caseSensitive, searchType):
		try:
			allText = get_full_text(textCtrl)
			if not allText or not allText.strip():
				log.debug("doQuickSearch: empty text from get_full_text")
				self.lastMatches = []
				self.lastSearchTerm = term
				self.lastCaseSensitive = caseSensitive
				self.lastSearchType = searchType
				self.newSearchPerformed = True
				return True  # search technically done, just no results

			searchFlags = 0 if caseSensitive else re.IGNORECASE
			isRegex = searchType == SearchType.REGULAR_EXPRESSION

			if isRegex:
				try:
					matches = re.finditer(term, allText, searchFlags)
				except re.error as e:
					log.error(f"Regex error: {e}")
					return False
			else:
				matches = re.finditer(re.escape(term), allText, searchFlags)

			self.lastMatches = []
			if term.lower() == "error":
				for m in matches:
					start = m.start()
					line_start = allText.rfind('\n', 0, start) + 1
					line_end = allText.find('\n', start)
					if line_end == -1:
						line_end = len(allText)
					line_text = allText[line_start:line_end]
					if not _is_excluded_error_line(line_text):
						self.lastMatches.append(m.span())
			else:
				self.lastMatches = [m.span() for m in matches]

			self.lastSearchTerm = term
			self.lastCaseSensitive = caseSensitive
			self.lastSearchType = searchType
			self.newSearchPerformed = True
			log.debug(f"Quick search for '{term}' found {len(self.lastMatches)} matches")
			return True
		except re.error as e:
			log.error(f"Error during quick search: {e}")
			return False

	def findNextMatch(self, caretPos, wrap):
		if not self.lastMatches:
			return -1
		for i, (start, end) in enumerate(self.lastMatches):
			if start > caretPos:
				return i
		if wrap:
			return 0
		return -1

	def findPrevMatch(self, caretPos, wrap):
		if not self.lastMatches:
			return -1
		for i in range(len(self.lastMatches)-1, -1, -1):
			start, end = self.lastMatches[i]
			if start < caretPos:
				return i
		if wrap:
			return len(self.lastMatches)-1
		return -1

	def moveToResult(self, textCtrl, index, announce_total=False):
		if not self.lastMatches or index < 0 or index >= len(self.lastMatches):
			ui.message(_("No matches available"))
			return

		start_pos, end_pos = self.lastMatches[index]
		line_num, line_text = extract_line_from_textctrl(textCtrl, start_pos)

		def _move():
			try:
				focusObj = api.getFocusObject()
				if not (focusObj and focusObj.role == controlTypes.Role.EDITABLETEXT and hasattr(focusObj, 'windowHandle') and focusObj.windowHandle == textCtrl.windowHandle):
					if hasattr(textCtrl, 'setFocus'):
						textCtrl.setFocus()
					else:
						api.setFocusObject(textCtrl)

				ti = textCtrl.makeTextInfo(textInfos.POSITION_ALL)
				ti.collapse()
				ti.move(textInfos.UNIT_CHARACTER, start_pos)
				ti.collapse()
				ti.updateSelection()

				core.callLater(50, self._speakResult, index, len(self.lastMatches), announce_total, line_num, line_text)
			except (RuntimeError, NotImplementedError) as e:
				log.error(f"Error moving to match: {e}")
				ui.message(_("Error moving to match"))

		wx.CallAfter(_move)

	# Hard cap on how much text is ever handed to ui.message() in one call.
	# SAPI5 and other synth drivers can lock up or crash NVDA outright when fed
	# a very long string in one shot (this is what caused the "string only"
	# mode to crash the synth and take NVDA down with it). One log line is
	# normally well under this, so the cap only bites on pathological lines.
	_MAX_SPOKEN_CHARS = 300

	def _speakResult(self, current_index, total_matches, announce_total, line_num, line_text):
		from .config_manager import PluginSettings
		try:
			announceMode = PluginSettings.get().announceMode
			parts = []
			if announceMode == "line_only":
				if announce_total:
					parts.append(_("Found {count} items").format(count=total_matches))
				parts.append(_("{line} line").format(line=line_num))
			elif announceMode == "string_only":
				safeLineText = line_text
				if len(safeLineText) > self._MAX_SPOKEN_CHARS:
					safeLineText = safeLineText[:self._MAX_SPOKEN_CHARS] + "..."
				parts.append(safeLineText)
			else:
				# "full" mode: term, line number, and match position together,
				# e.g. "error line 125 1 of 30" -- the line number was
				# previously missing from this format entirely.
				if announce_total:
					parts.append(_("Found {count} items").format(count=total_matches))
				parts.append(_("{term} line {line} {current} of {total}").format(
					term=self.lastSearchTerm,
					line=line_num,
					current=current_index + 1,
					total=total_matches
				))
			full_message = " ".join(parts)
			if len(full_message) > self._MAX_SPOKEN_CHARS:
				full_message = full_message[:self._MAX_SPOKEN_CHARS] + "..."
			ui.message(full_message)
		except LookupError as e:
			log.error(f"Error speaking result: {e}")
			ui.message(_("{current} of {total}").format(current=current_index + 1, total=total_matches))


class LogSearchDialog(wx.Dialog):
	def __init__(self, parent, logTextCtrl, globalPluginInstance):
		from .config_manager import SearchHistory
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
		self._pendingFocusMove = False

		self.panel = wx.Panel(self)
		self.mainSizer = wx.BoxSizer(wx.VERTICAL)

		searchSizer = wx.BoxSizer(wx.HORIZONTAL)
		self.searchBox = wx.ComboBox(self.panel, style=wx.CB_DROPDOWN | wx.TE_PROCESS_ENTER, choices=self.searchHistory.getItems())
		self.searchBox.SetValue("error")
		searchSizer.Add(self.searchBox, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)

		self.searchButton = wx.Button(self.panel, label=_("Find Next"))
		searchSizer.Add(self.searchButton, flag=wx.ALIGN_CENTER_VERTICAL | wx.TOP | wx.BOTTOM | wx.RIGHT, border=5)

		self.findAndFocusButton = wx.Button(self.panel, label=_("Find & Focus"))
		searchSizer.Add(self.findAndFocusButton, flag=wx.ALIGN_CENTER_VERTICAL | wx.TOP | wx.BOTTOM | wx.RIGHT, border=5)

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
		optionsSizer.Add(self.searchTypeCombo, flag=wx.ALL | wx.EXPAND, border=5)
		self.mainSizer.Add(optionsSizer, flag=wx.EXPAND)

		self.resultBox = wx.TextCtrl(self.panel, style=wx.TE_MULTILINE | wx.TE_READONLY)
		self.mainSizer.Add(self.resultBox, proportion=1, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, border=5)

		buttonSizer = wx.BoxSizer(wx.HORIZONTAL)
		self.prevButton = wx.Button(self.panel, label=_("Find Previous"))
		buttonSizer.Add(self.prevButton, flag=wx.ALL, border=5)

		self.cancelButton = wx.Button(self.panel, id=wx.ID_CANCEL, label=_("Close"))
		buttonSizer.Add(self.cancelButton, flag=wx.ALL, border=5)
		self.mainSizer.Add(buttonSizer, flag=wx.ALIGN_RIGHT | wx.ALL, border=5)

		self.panel.SetSizer(self.mainSizer)

		self.searchButton.Bind(wx.EVT_BUTTON, lambda evt: self.performSearch(forward=True, focus=False))
		self.findAndFocusButton.Bind(wx.EVT_BUTTON, lambda evt: self.performSearch(forward=True, focus=True))
		self.searchBox.Bind(wx.EVT_TEXT_ENTER, lambda evt: self.performSearch(forward=True, focus=True))
		self.prevButton.Bind(wx.EVT_BUTTON, lambda evt: self.performSearch(forward=False, focus=False))
		self.cancelButton.Bind(wx.EVT_BUTTON, self.onClose)
		self.Bind(wx.EVT_CLOSE, self.onClose)
		# Event-driven focus handoff: rather than guessing how long window
		# destruction takes with a hardcoded core.callLater delay, wait for the
		# actual OS destroy notification before touching the caret/focus.
		self.Bind(wx.EVT_WINDOW_DESTROY, self._onWindowDestroy)

		self.searchBox.SetFocus()

	def onClose(self, event):
		self.dialogOpen = False
		if self.globalPlugin and self.globalPlugin.searchDialog is self:
			self.globalPlugin.searchDialog = None
		self.Destroy()

	def _onWindowDestroy(self, event):
		event.Skip()
		if event.GetEventObject() is not self:
			return
		if self._pendingFocusMove:
			self._pendingFocusMove = False
			wx.CallAfter(self._focusAndMoveToMatch)

	def doSearch(self, term, caseSensitive, searchType):
		"""Runs entirely off the main thread (called from a worker thread by
		performSearch). Retrieves the full log text and performs the regex scan,
		then returns a plain list of match spans -- never touches wx widgets or
		self.matches directly, so the caller must marshal the result back to the
		main thread via wx.CallAfter before applying it to UI state."""
		with self.searchLock:
			try:
				allText = get_full_text(self.logCtrl)
				if not allText or not allText.strip():
					return []
				searchFlags = 0 if caseSensitive else re.IGNORECASE
				isRegex = searchType == SearchType.REGULAR_EXPRESSION
				if isRegex:
					try:
						found_iter = re.finditer(term, allText, searchFlags)
					except re.error as e:
						log.error(f"Invalid regular expression: {e}")
						return None
				else:
					found_iter = re.finditer(re.escape(term), allText, searchFlags)

				matches = []
				if term.lower() == "error":
					for m in found_iter:
						start = m.start()
						line_start = allText.rfind('\n', 0, start) + 1
						line_end = allText.find('\n', start)
						if line_end == -1:
							line_end = len(allText)
						line_text = allText[line_start:line_end]
						if not _is_excluded_error_line(line_text):
							matches.append(m.span())
				else:
					matches = [m.span() for m in found_iter]
				return matches
			except re.error as e:
				log.error(f"Error during search: {e}")
				return None

	def getCaretPosition(self):
		try:
			textInfo = self.logCtrl.makeTextInfo(textInfos.POSITION_CARET)
			return textInfo.bookmark.startOffset
		except (RuntimeError, NotImplementedError) as e:
			log.error(f"Error getting caret position: {e}")
			return 0

	def performSearch(self, forward=True, focus=False):
		if not self.dialogOpen:
			return

		term = self.searchBox.GetValue().strip()
		if not term:
			ui.message(_("Search term cannot be empty"))
			return
		self.searchHistory.append(term)
		caseSensitive = self.caseSensitiveCheck.GetValue()
		wrap = self.wrapCheck.GetValue()
		searchType = SearchType.getByIndex(self.searchTypeCombo.GetSelection())
		config.conf["LogViewerPlugin"]["searchCaseSensitivity"] = caseSensitive
		config.conf["LogViewerPlugin"]["searchWrap"] = wrap
		config.conf["LogViewerPlugin"]["searchType"] = searchType.name
		config.conf.save()

		needsNewSearch = (
			term != self.lastSearchTerm or
			caseSensitive != self.lastCaseSensitive or
			searchType != self.lastSearchType or
			not self.matches
		)

		if not needsNewSearch:
			self._continueSearchNavigation(wrap, forward, focus)
			return

		# Offload the full-text retrieval and regex scan to a background thread
		# so a large log file cannot trip NVDA's main-thread watchdog timeout.
		self.resultBox.SetValue(_("Searching..."))

		def worker():
			matches = self.doSearch(term, caseSensitive, searchType)
			wx.CallAfter(self._onSearchComplete, term, caseSensitive, searchType, wrap, matches, forward, focus)

		threading.Thread(target=worker, daemon=True).start()

	def _onSearchComplete(self, term, caseSensitive, searchType, wrap, matches, forward, focus):
		if not self.dialogOpen:
			return
		if matches is None:
			self.resultBox.SetValue(_("Search failed or invalid expression"))
			ui.message(_("No matches found"))
			return
		self.matches = matches
		self.lastSearchTerm = term
		self.lastCaseSensitive = caseSensitive
		self.lastSearchType = searchType
		self._continueSearchNavigation(wrap, forward, focus)

	def _continueSearchNavigation(self, wrap, forward, focus):
		if not self.dialogOpen:
			return
		if not self.matches:
			self.resultBox.SetValue(_("No matches found"))
			ui.message(_("No matches found"))
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
				ui.message(_("Wrapping to first match"))
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
				ui.message(_("Wrapping to last match"))
				found = True

		if not found:
			ui.message(_("No matches found"))
			return

		if self.globalPlugin:
			self.globalPlugin.search_manager.lastSearchTerm = self.lastSearchTerm
			self.globalPlugin.search_manager.lastMatches = self.matches
			self.globalPlugin.search_manager.currentMatchIndex = self.currentMatch
			self.globalPlugin.search_manager.lastCaseSensitive = self.lastCaseSensitive
			self.globalPlugin.search_manager.lastSearchType = self.lastSearchType
			self.globalPlugin.search_manager.newSearchPerformed = True

			self.globalPlugin.lastSearchTerm = self.lastSearchTerm
			self.globalPlugin.lastMatches = self.matches
			self.globalPlugin.currentMatchIndex = self.currentMatch
			self.globalPlugin._lastSearchCaseSensitive = self.lastCaseSensitive
			self.globalPlugin._lastSearchType = self.lastSearchType

		if not focus:
			self.updateResultDisplay()

		if focus:
			# Defer the actual focus move until wx confirms the dialog window has
			# genuinely finished being destroyed (see _onWindowDestroy), instead of
			# guessing at a fixed millisecond delay.
			self._pendingFocusMove = True
			self.Destroy()
		else:
			self.moveToMatch()

		ui.message(_("Found {count} matches.").format(count=len(self.matches)))

	def _focusAndMoveToMatch(self):
		if not self.matches or self.currentMatch < 0 or self.currentMatch >= len(self.matches):
			ui.message(_("No matches available"))
			return
		start_pos, end_pos = self.matches[self.currentMatch]

		def _move():
			try:
				focusObj = api.getFocusObject()
				if not (focusObj and focusObj.role == controlTypes.Role.EDITABLETEXT and hasattr(focusObj, 'windowHandle') and focusObj.windowHandle == self.logCtrl.windowHandle):
					if hasattr(self.logCtrl, 'setFocus'):
						self.logCtrl.setFocus()
					else:
						api.setFocusObject(self.logCtrl)

				ti = self.logCtrl.makeTextInfo(textInfos.POSITION_ALL)
				ti.collapse()
				ti.move(textInfos.UNIT_CHARACTER, start_pos)
				ti.collapse()
				ti.updateSelection()

				line_num = ti.text.count('\n', 0, start_pos) + 1
				line_start = ti.text.rfind('\n', 0, start_pos) + 1
				line_end = ti.text.find('\n', end_pos)
				if line_end == -1:
					line_end = len(ti.text)
				line_text = ti.text[line_start:line_end].strip()
				ui.message(_("Line {number}: {text}").format(number=line_num, text=line_text))
			except (RuntimeError, NotImplementedError) as e:
				log.error(f"Error moving to match: {e}")
				ui.message(_("Error moving to match"))

		wx.CallAfter(_move)

	def updateResultDisplay(self):
		if not self.dialogOpen:
			return

		displayText = []
		if not self.matches:
			self.resultBox.SetValue(_("No matches found"))
			return
		start_idx = max(0, self.currentMatch - 2)
		end_idx = min(len(self.matches), self.currentMatch + 3)
		allText = get_full_text(self.logCtrl)
		if not allText:
			return
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

	def moveToMatch(self):
		if not self.dialogOpen:
			return
		if not self.matches or self.currentMatch < 0 or self.currentMatch >= len(self.matches):
			ui.message(_("No matches available"))
			return
		start_pos, end_pos = self.matches[self.currentMatch]

		def _move():
			try:
				ti = self.logCtrl.makeTextInfo(textInfos.POSITION_ALL)
				ti.collapse()
				ti.move(textInfos.UNIT_CHARACTER, start_pos)
				ti.collapse()
				ti.updateSelection()

				line_num = ti.text.count('\n', 0, start_pos) + 1
				line_start = ti.text.rfind('\n', 0, start_pos) + 1
				line_end = ti.text.find('\n', end_pos)
				if line_end == -1:
					line_end = len(ti.text)
				line_text = ti.text[line_start:line_end].strip()
				ui.message(_("Line {number}: {text}").format(number=line_num, text=line_text))
			except (RuntimeError, NotImplementedError) as e:
				log.error(f"Error moving to match: {e}")
				ui.message(_("Error moving to match"))

		wx.CallAfter(_move)

	def isNVDAViewerObject(self, obj):
		try:
			return obj.role == controlTypes.Role.EDITABLETEXT and fIsLogViewer(obj)
		except (AttributeError, RuntimeError):
			return False
