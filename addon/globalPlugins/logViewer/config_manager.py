# config_manager.py

import json
import os
import shutil
from logHandler import log
import config
import globalVars
import addonHandler

addonHandler.initTranslation()

DEFAULT_TERMS = ["error", "warning", "debug"]
DEFAULT_ANNOUNCE_MODE = "full"
DEFAULT_QUICK_SEARCH_TERMS = ["error", "warning", "DEBUGWARNING"]
DEFAULT_SEARCH_TYPE = "NORMAL"
VALID_SEARCH_TYPES = ("NORMAL", "REGULAR_EXPRESSION")
HISTORY_LIMIT = 20

LEGACY_INI_SECTION = "LogViewerPlugin"
LEGACY_INI_SEARCH_KEYS = ("searchCaseSensitivity", "searchWrap", "searchType", "searchHistory")

ANNOUNCE_MODE_CHOICES = (
	("full", _("Term, line number and match position (e.g. \"error line 125 1 of 30\")")),
	("line_only", _("Line number only (e.g. \"125 line\")")),
	("string_only", _("Matched line text only")),
)


def get_addon_dir():
	"""Returns userConfig\\ChaiChaimee\\logViewer, creating it if necessary."""
	addonDir = os.path.join(globalVars.appArgs.configPath, "ChaiChaimee", "logViewer")
	if not os.path.isdir(addonDir):
		try:
			os.makedirs(addonDir)
		except OSError as e:
			log.error(f"Could not create directory {addonDir}: {e}")
	return addonDir


def get_history_file_path():
	return os.path.join(get_addon_dir(), "logViewer.json")


def get_settings_file_path():
	return os.path.join(get_addon_dir(), "settings.json")


def _readJson(filePath):
	"""Returns (data, usedPath). Falls back to the rolling backup when the main file is unreadable."""
	for candidate in (filePath, filePath + ".bak"):
		if not os.path.isfile(candidate):
			continue
		try:
			with open(candidate, "r", encoding="utf-8") as jsonFile:
				return json.load(jsonFile), candidate
		except (OSError, ValueError) as e:
			log.error(f"Could not read {candidate}: {e}")
	return None, None


def _writeJsonAtomic(filePath, payload):
	tempPath = filePath + ".tmp"
	backupPath = filePath + ".bak"
	try:
		os.makedirs(os.path.dirname(filePath), exist_ok=True)
		with open(tempPath, "w", encoding="utf-8") as tempFile:
			json.dump(payload, tempFile, ensure_ascii=False, indent=2)
			tempFile.flush()
			os.fsync(tempFile.fileno())
		if os.path.isfile(filePath):
			shutil.copy2(filePath, backupPath)
		os.replace(tempPath, filePath)
		return True
	except (OSError, TypeError, ValueError) as e:
		log.error(f"Could not write {filePath}: {e}")
		if os.path.isfile(tempPath):
			try:
				os.remove(tempPath)
			except OSError as cleanupError:
				log.error(f"Could not remove temporary file {tempPath}: {cleanupError}")
		return False


def _isInsideConfigPath(filePath):
	configRoot = os.path.abspath(globalVars.appArgs.configPath)
	try:
		return os.path.commonpath([configRoot, os.path.abspath(filePath)]) == configRoot
	except ValueError:
		return False


def _removeLegacyFile(filePath):
	if not _isInsideConfigPath(filePath):
		log.error(f"Refusing to delete {filePath}: outside the NVDA configuration folder")
		return
	if os.path.isfile(filePath):
		try:
			os.remove(filePath)
			log.info(f"Removed migrated legacy file {filePath}")
		except OSError as e:
			log.error(f"Could not remove legacy file {filePath}: {e}")


def _validTermList(data):
	if not isinstance(data, list) or not all(isinstance(term, str) for term in data):
		return None
	terms = [term for term in data if term.strip()][:HISTORY_LIMIT]
	return terms or None


def _toBool(value, default):
	if isinstance(value, bool):
		return value
	if isinstance(value, str):
		lowered = value.strip().lower()
		if lowered in ("true", "1", "yes", "on"):
			return True
		if lowered in ("false", "0", "no", "off"):
			return False
	return default


class SearchHistory:
	_instance = None

	@classmethod
	def get(cls):
		if cls._instance is None:
			cls._instance = cls()
		return cls._instance

	def __init__(self):
		self._terms = []
		self._historyFile = get_history_file_path()
		self._importLegacyFile()
		self._load()
		log.debug(f"Search history ready: {self._historyFile}, {len(self._terms)} terms")

	def _importLegacyFile(self):
		if os.path.isfile(self._historyFile):
			return
		configRoot = globalVars.appArgs.configPath
		legacyFiles = (
			os.path.join(configRoot, "ChaiChaimee", "logViewer.json"),
			os.path.join(configRoot, "logViewer.json"),
		)
		for legacyFile in legacyFiles:
			if not os.path.isfile(legacyFile):
				continue
			terms = _validTermList(_readJson(legacyFile)[0])
			if terms is None:
				log.error(f"Legacy search history {legacyFile} holds no usable data, ignoring it")
				continue
			self._terms = terms
			if not self.save():
				return
			log.info(f"Migrated search history from {legacyFile} to {self._historyFile}")
			_removeLegacyFile(legacyFile)
			return

	def _load(self):
		data, usedPath = _readJson(self._historyFile)
		terms = _validTermList(data)
		if terms is not None:
			self._terms = terms
			return
		if data is not None:
			log.error(f"Search history in {usedPath} is unusable, resetting to defaults")
		if not self._terms:
			self._terms = DEFAULT_TERMS.copy()
		self.save()

	def save(self):
		return _writeJsonAtomic(self._historyFile, self._terms)

	def getItems(self):
		return list(self._terms)

	def append(self, term):
		if not term:
			return
		loweredTerm = term.lower()
		self._terms = [existing for existing in self._terms if existing.lower() != loweredTerm]
		self._terms.insert(0, term)
		del self._terms[HISTORY_LIMIT:]
		self.save()

	def addOlderTerms(self, terms):
		knownTerms = {existing.lower() for existing in self._terms}
		for term in terms:
			if len(self._terms) >= HISTORY_LIMIT:
				break
			if term.strip() and term.lower() not in knownTerms:
				self._terms.append(term)
				knownTerms.add(term.lower())


class PluginSettings:
	"""User preferences stored as JSON in userConfig\\ChaiChaimee\\logViewer\\settings.json,
	next to the search history file."""

	_instance = None

	@classmethod
	def get(cls):
		if cls._instance is None:
			cls._instance = cls()
		return cls._instance

	def __init__(self):
		self._settingsFile = get_settings_file_path()
		self.announceMode = DEFAULT_ANNOUNCE_MODE
		self.quickSearchTerms = DEFAULT_QUICK_SEARCH_TERMS.copy()
		self.searchCaseSensitive = False
		self.searchWrap = True
		self.searchType = DEFAULT_SEARCH_TYPE
		self.hasSearchOptions = False
		self.load()

	def load(self):
		data, usedPath = _readJson(self._settingsFile)
		if data is None:
			return
		if not isinstance(data, dict):
			log.error(f"Settings in {usedPath} are unusable, using defaults")
			return
		announceMode = data.get("announceMode")
		if announceMode in [choice[0] for choice in ANNOUNCE_MODE_CHOICES]:
			self.announceMode = announceMode
		terms = data.get("quickSearchTerms")
		if isinstance(terms, list) and terms and all(isinstance(term, str) for term in terms):
			self.quickSearchTerms = terms
		self.searchCaseSensitive = _toBool(data.get("searchCaseSensitive"), self.searchCaseSensitive)
		self.searchWrap = _toBool(data.get("searchWrap"), self.searchWrap)
		if data.get("searchType") in VALID_SEARCH_TYPES:
			self.searchType = data["searchType"]
		self.hasSearchOptions = all(
			key in data for key in ("searchCaseSensitive", "searchWrap", "searchType")
		)

	def save(self):
		saved = _writeJsonAtomic(
			self._settingsFile,
			{
				"announceMode": self.announceMode,
				"quickSearchTerms": self.quickSearchTerms,
				"searchCaseSensitive": self.searchCaseSensitive,
				"searchWrap": self.searchWrap,
				"searchType": self.searchType,
			},
		)
		if saved:
			self.hasSearchOptions = True
		return saved

	def ensureFileExists(self):
		if not os.path.isfile(self._settingsFile):
			self.save()

	def setSearchOptions(self, caseSensitive, wrap, searchTypeName):
		self.searchCaseSensitive = bool(caseSensitive)
		self.searchWrap = bool(wrap)
		self.searchType = searchTypeName if searchTypeName in VALID_SEARCH_TYPES else DEFAULT_SEARCH_TYPE
		return self.save()


def _getBaseIniSection():
	try:
		return config.conf.profiles[0][LEGACY_INI_SECTION]
	except (AttributeError, IndexError, KeyError):
		return None


def _parseLegacyHistory(rawValue):
	if not isinstance(rawValue, str) or rawValue in ("", "[]"):
		return []
	try:
		terms = json.loads(rawValue)
	except ValueError as e:
		log.error(f"Legacy searchHistory in nvda.ini is not valid JSON: {e}")
		return []
	return _validTermList(terms) or []


def importLegacyNvdaIniSettings():
	"""Moves search options and history out of nvda.ini into the add-on's own JSON files.

	Idempotent: when none of the legacy keys exist it does nothing. The keys are deleted
	from nvda.ini only after both JSON files were written successfully."""
	section = _getBaseIniSection()
	if section is None:
		return
	foundKeys = [key for key in LEGACY_INI_SEARCH_KEYS if key in section]
	if not foundKeys:
		return
	log.info(f"Found legacy search settings in nvda.ini: {', '.join(foundKeys)}")
	settings = PluginSettings.get()
	history = SearchHistory.get()
	if not settings.hasSearchOptions:
		settings.searchCaseSensitive = _toBool(section.get("searchCaseSensitivity"), False)
		settings.searchWrap = _toBool(section.get("searchWrap"), True)
		legacyType = section.get("searchType", DEFAULT_SEARCH_TYPE)
		settings.searchType = legacyType if legacyType in VALID_SEARCH_TYPES else DEFAULT_SEARCH_TYPE
	history.addOlderTerms(_parseLegacyHistory(section.get("searchHistory")))
	if not (settings.save() and history.save()):
		log.error("Legacy search settings could not be stored in JSON; nvda.ini was left untouched")
		return
	for key in foundKeys:
		del section[key]
	try:
		config.conf.save()
	except (OSError, ValueError) as e:
		log.error(f"Could not save nvda.ini after removing legacy search settings: {e}")
		return
	log.info("Legacy search settings moved from nvda.ini to " + get_addon_dir())


def initConfiguration():
	confspec = {
		"bookmarkCount": "integer(default=1)",
		"lastBootTimestamp": "float(default=0.0)",
	}
	config.conf.spec[LEGACY_INI_SECTION] = confspec
