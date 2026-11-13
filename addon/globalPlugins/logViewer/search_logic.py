# search_logic.py

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
from bisect import bisect_left, bisect_right
from collections import namedtuple
from enum import Enum, unique
import threading
import time
import ctypes
from ctypes import wintypes

from .config_manager import PluginSettings

addonHandler.initTranslation()

GA_PARENT = 1
user32 = ctypes.windll.user32
user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
user32.GetAncestor.restype = wintypes.HWND

MAX_PATTERN_LENGTH = 500
SEARCH_TIME_BUDGET_SECONDS = 3.0
MAX_MATCHES = 200000

SearchRequest = namedtuple("SearchRequest", ("term", "caseSensitive", "wrap", "searchType"))


@unique
class SearchType(Enum):
	NORMAL = "normal"
	REGULAR_EXPRESSION = "regular expression"

	@staticmethod
	def getByIndex(index):
		return list(SearchType)[index]

	@staticmethod
	def getIndexByName(name):
		for index, searchType in enumerate(SearchType):
			if searchType.name == name:
				return index
		return 0

	@staticmethod
	def getByName(name):
		for searchType in SearchType:
			if searchType.name == name:
				return searchType
		return SearchType.NORMAL

	@staticmethod
	def getSearchTypes():
		return [_(searchType.value) for searchType in SearchType]


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
	"""Must be called on the main thread: it reads a wx control and a COM-backed text info."""
	if fIsLogViewer(textCtrl):
		try:
			viewer = gui.logViewer.logViewer
			outputCtrl = getattr(viewer, 'outputCtrl', None)
			if viewer is not None and outputCtrl is not None:
				text = outputCtrl.GetValue()
				if text.strip():
					return text
		except (AttributeError, RuntimeError) as e:
			log.error(f"outputCtrl.GetValue failed: {e}, falling back to makeTextInfo")
	try:
		return textCtrl.makeTextInfo(textInfos.POSITION_ALL).text
	except (RuntimeError, NotImplementedError) as e:
		log.error(f"Error getting text from text control: {e}")
		return ""


def getLineInfo(text, position):
	if not text:
		return 0, ""
	lineNumber = text.count('\n', 0, position) + 1
	lineStart = text.rfind('\n', 0, position) + 1
	lineEnd = text.find('\n', position)
	if lineEnd == -1:
		lineEnd = len(text)
	return lineNumber, text[lineStart:lineEnd].strip()


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


_UNBOUNDED_TAIL = re.compile(r"(?<!\\)(?:[+*]|\{\d*,\d*\})\??$")
_GROUP_PREFIX = re.compile(r"^\?(?:P<[^>]+>|<[^>=!][^>]*>|[:=!]|<[=!])")
_QUANTIFIER_AT = re.compile(r"(?:\*|\+|\?|\{\d*,?\d*\})\??")
_PLAIN_LITERAL = re.compile(r"[^\\\[\](){}.*+?^$|]*")


def _iterGroups(pattern):
	"""Yields (groupContent, quantifierAfterGroup) for every closed group of a pattern."""
	openStarts = []
	index = 0
	inClass = False
	length = len(pattern)
	while index < length:
		char = pattern[index]
		if char == "\\":
			index += 2
			continue
		if inClass:
			if char == "]":
				inClass = False
			index += 1
			continue
		if char == "[":
			inClass = True
			index += 1
			if index < length and pattern[index] == "^":
				index += 1
			if index < length and pattern[index] == "]":
				index += 1
			continue
		if char == "(":
			openStarts.append(index + 1)
		elif char == ")" and openStarts:
			start = openStarts.pop()
			quantifier = _QUANTIFIER_AT.match(pattern, index + 1)
			yield pattern[start:index], quantifier.group(0) if quantifier else ""
		index += 1


def _splitTopLevelAlternatives(groupContent):
	branches = []
	depth = 0
	inClass = False
	branchStart = 0
	index = 0
	while index < len(groupContent):
		char = groupContent[index]
		if char == "\\":
			index += 2
			continue
		if inClass:
			inClass = char != "]"
		elif char == "[":
			inClass = True
		elif char == "(":
			depth += 1
		elif char == ")":
			depth -= 1
		elif char == "|" and depth == 0:
			branches.append(groupContent[branchStart:index])
			branchStart = index + 1
		index += 1
	branches.append(groupContent[branchStart:])
	return branches


def _isUnboundedQuantifier(quantifier):
	if quantifier.startswith(("*", "+")):
		return True
	if quantifier.startswith("{") and "," in quantifier:
		upperBound = quantifier.strip("{}?").split(",")[1]
		return not upperBound or int(upperBound) > 10
	return False


def _hasOverlappingLiteralBranches(branches):
	if len(branches) < 2 or not all(_PLAIN_LITERAL.fullmatch(branch) for branch in branches):
		return False
	for first in range(len(branches)):
		for second in range(len(branches)):
			if first == second:
				continue
			shorter, longer = branches[first], branches[second]
			if shorter and longer.startswith(shorter):
				return True
	return False


def isRiskyPattern(pattern):
	"""Heuristic screen for shapes that cause catastrophic backtracking: an unbounded repeat of a
	group that itself ends in an unbounded repeat, or of alternatives where one is a prefix of another.
	Python's re module cannot be interrupted, so a pattern that gets this far must not be run."""
	for groupContent, quantifier in _iterGroups(pattern):
		if not _isUnboundedQuantifier(quantifier):
			continue
		content = _GROUP_PREFIX.sub("", groupContent, count=1)
		if _UNBOUNDED_TAIL.search(content):
			return True
		if _hasOverlappingLiteralBranches(_splitTopLevelAlternatives(content)):
			return True
	return False


def compileSearchPattern(term, caseSensitive, searchType):
	"""Returns (compiledPattern, status). status is ok, invalid, risky or tooLong."""
	flags = 0 if caseSensitive else re.IGNORECASE
	if searchType != SearchType.REGULAR_EXPRESSION:
		return re.compile(re.escape(term), flags), "ok"
	if len(term) > MAX_PATTERN_LENGTH:
		return None, "tooLong"
	if isRiskyPattern(term):
		return None, "risky"
	try:
		return re.compile(term, flags), "ok"
	except re.error as e:
		log.error(f"Invalid regular expression: {e}")
		return None, "invalid"


def describeSearchFailure(status):
	if status == "invalid":
		return _("Invalid regular expression")
	if status == "risky":
		return _("This regular expression can be very slow. Remove nested repeats such as (a+)+")
	if status == "tooLong":
		return _("Regular expression is too long")
	if status == "timeout":
		return _("Search aborted because it took too long")
	if status == "tooMany":
		return _("Too many matches. Refine the search")
	return _("Search failed")


def collectMatchSpans(compiledPattern, allText, skipSpellingNoise):
	"""Runs on a worker thread. Returns (spans, status); status is ok, timeout or tooMany."""
	spans = []
	deadline = time.monotonic() + SEARCH_TIME_BUDGET_SECONDS
	for match in compiledPattern.finditer(allText):
		if match.end() == match.start():
			continue
		if skipSpellingNoise:
			start = match.start()
			lineStart = allText.rfind('\n', 0, start) + 1
			lineEnd = allText.find('\n', start)
			if lineEnd == -1:
				lineEnd = len(allText)
			if _is_excluded_error_line(allText[lineStart:lineEnd]):
				continue
		spans.append(match.span())
		if len(spans) > MAX_MATCHES:
			return None, "tooMany"
		if time.monotonic() > deadline:
			return None, "timeout"
	return spans, "ok"


class SearchManager:
	_MAX_SPOKEN_CHARS = 300

	def __init__(self):
		self.lastSearchTerm = ""
		self.lastMatches = []
		self.currentMatchIndex = -1
		self.lastCaseSensitive = None
		self.lastSearchType = None
		self.newSearchPerformed = False
		self._matchStarts = []
		self._lastText = ""
		self._generation = 0

	@property
	def lastText(self):
		return self._lastText

	def matchesQuery(self, term, caseSensitive, searchType):
		return (
			self.lastSearchTerm == term
			and self.lastCaseSensitive == caseSensitive
			and self.lastSearchType == searchType
		)

	def invalidate(self):
		self._generation += 1
		self.lastMatches = []
		self._matchStarts = []
		self._lastText = ""
		self.currentMatchIndex = -1

	def _store(self, term, caseSensitive, searchType, allText, spans):
		self.lastSearchTerm = term
		self.lastCaseSensitive = caseSensitive
		self.lastSearchType = searchType
		self.lastMatches = spans
		self._matchStarts = [span[0] for span in spans]
		self._lastText = allText
		self.currentMatchIndex = -1
		self.newSearchPerformed = True

	def requestMatches(self, textCtrl, term, caseSensitive, searchType, onResult):
		"""Main thread only. The text is read here; the scan runs on a worker thread and
		onResult(status) is called on the main thread once it finished."""
		allText = get_full_text(textCtrl)
		if not allText or not allText.strip():
			self._generation += 1
			self._store(term, caseSensitive, searchType, "", [])
			onResult("ok")
			return
		if self.matchesQuery(term, caseSensitive, searchType) and allText == self._lastText:
			onResult("ok")
			return
		compiledPattern, patternStatus = compileSearchPattern(term, caseSensitive, searchType)
		if compiledPattern is None:
			onResult(patternStatus)
			return
		self._generation += 1
		generation = self._generation
		skipSpellingNoise = term.lower() == "error"

		def worker():
			spans, scanStatus = collectMatchSpans(compiledPattern, allText, skipSpellingNoise)
			wx.CallAfter(self._onScanFinished, generation, term, caseSensitive, searchType, allText, spans, scanStatus, onResult)

		threading.Thread(target=worker, daemon=True).start()

	def _onScanFinished(self, generation, term, caseSensitive, searchType, allText, spans, scanStatus, onResult):
		if generation != self._generation:
			return
		if scanStatus != "ok":
			onResult(scanStatus)
			return
		self._store(term, caseSensitive, searchType, allText, spans)
		log.debug(f"Search for {term!r} found {len(spans)} matches")
		onResult("ok")

	def findNextMatch(self, caretPos, wrap, inclusive=False):
		if not self.lastMatches:
			return -1
		if inclusive:
			index = bisect_left(self._matchStarts, caretPos)
		else:
			index = bisect_right(self._matchStarts, caretPos)
		if index < len(self.lastMatches):
			return index
		return 0 if wrap else -1

	def findPrevMatch(self, caretPos, wrap):
		if not self.lastMatches:
			return -1
		index = bisect_left(self._matchStarts, caretPos) - 1
		if index >= 0:
			return index
		return len(self.lastMatches) - 1 if wrap else -1

	def moveToResult(self, textCtrl, index, announce_total=False):
		if not self.lastMatches or index < 0 or index >= len(self.lastMatches):
			ui.message(_("No matches available"))
			return

		start_pos = self.lastMatches[index][0]
		line_num, line_text = getLineInfo(self._lastText, start_pos)
		totalMatches = len(self.lastMatches)

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

				core.callLater(50, self._speakResult, index, totalMatches, announce_total, line_num, line_text)
			except (RuntimeError, NotImplementedError) as e:
				log.error(f"Error moving to match: {e}")
				ui.message(_("Error moving to match"))

		wx.CallAfter(_move)

	def _speakResult(self, current_index, total_matches, announce_total, line_num, line_text):
		try:
			announceMode = PluginSettings.get().announceMode
			parts = []
			if announceMode == "line_only":
				if announce_total:
					parts.append(_("Found {count} items").format(count=total_matches))
				parts.append(_("{line} line").format(line=line_num))
			elif announceMode == "string_only":
				parts.append(line_text)
			else:
				if announce_total:
					parts.append(_("Found {count} items").format(count=total_matches))
				parts.append(_("{term} line {line} {current} of {total}").format(
					term=self.lastSearchTerm,
					line=line_num,
					current=current_index + 1,
					total=total_matches
				))
			full_message = " ".join(parts)
			# One capped string feeds both speech and braille; long lines can hang some synth drivers.
			if len(full_message) > self._MAX_SPOKEN_CHARS:
				full_message = full_message[:self._MAX_SPOKEN_CHARS] + _("... truncated")
			ui.message(full_message)
		except LookupError as e:
			log.error(f"Error speaking result: {e}")
			ui.message(_("{current} of {total}").format(current=current_index + 1, total=total_matches))
