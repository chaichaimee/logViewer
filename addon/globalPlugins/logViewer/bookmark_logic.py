# bookmark_logic.py

import re

# A bookmark is a line of its own. Matching "BOOKMARK n" anywhere inside a line
# also hit log lines that merely quote a bookmark and cut copied segments short.
BOOKMARK_LINE_PATTERN = re.compile(r"^BOOKMARK (\d+)[ \t\r]*$", re.MULTILINE)


def findBookmarks(logText):
	return [
		(match.start(), match.end(), int(match.group(1)))
		for match in BOOKMARK_LINE_PATTERN.finditer(logText)
	]


def pickLogPath(preferOld, currentPath, oldPath, pathExists):
	ordered = (oldPath, currentPath) if preferOld else (currentPath, oldPath)
	for candidate in ordered:
		if candidate and pathExists(candidate):
			return candidate
	return None


def pickCursorIndex(bookmarkCount, currentCursor, stepBack):
	"""Returns (index, wrapped) or (None, False) when there is no bookmark."""
	if bookmarkCount <= 0:
		return None, False
	latest = bookmarkCount - 1
	if not stepBack:
		return latest, False
	baseCursor = latest if currentCursor is None or currentCursor >= bookmarkCount else currentCursor
	newCursor = baseCursor - 1
	if newCursor < 0:
		return latest, True
	return newCursor, False


def extractSegment(logText, bookmarks, index):
	if index < 0 or index >= len(bookmarks):
		return None
	segmentStart = bookmarks[index][1]
	if index + 1 < len(bookmarks):
		segmentEnd = bookmarks[index + 1][0]
	else:
		segmentEnd = len(logText)
	return logText[segmentStart:segmentEnd].strip("\r\n")


def countLines(segmentText):
	if not segmentText:
		return 0
	return segmentText.count("\n") + 1
