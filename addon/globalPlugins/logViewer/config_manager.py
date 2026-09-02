# config_manager.py
# Copyright (C) 2026 Chai Chaimee
# Licensed under GNU General Public License. See COPYING.txt for details.

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

# Human readable labels are resolved lazily via _() so translation is applied
# whenever the settings panel actually reads this list, not at import time.
ANNOUNCE_MODE_CHOICES = (
	("full", _("Term, line number and match position (e.g. \"error line 125 1 of 30\")")),
	("line_only", _("Line number only (e.g. \"125 line\")")),
	("string_only", _("Matched line text only")),
)


def get_addon_dir():
	"""Returns userConfig\\ChaiChaimee\\logViewer, creating it if necessary."""
	base_dir = globalVars.appArgs.configPath
	addon_dir = os.path.join(base_dir, "ChaiChaimee", "logViewer")
	if not os.path.exists(addon_dir):
		try:
			os.makedirs(addon_dir)
		except OSError as e:
			log.error(f"Could not create directory {addon_dir}: {e}")
	return addon_dir


def get_history_file_path():
	return os.path.join(get_addon_dir(), "logViewer.json")


def get_settings_file_path():
	return os.path.join(get_addon_dir(), "settings.json")


class SearchHistory:
	_instance = None

	@classmethod
	def get(cls):
		if cls._instance is None:
			cls._instance = cls()
		return cls._instance

	def __init__(self):
		self._terms = []
		self._history_file = get_history_file_path()

		self._migrate_from_legacy_root_file()
		self._migrate_from_previous_chai_dir()
		self._migrate_from_config()
		self.load()  # Load existing history from new file

		if not self._terms:
			self._terms = DEFAULT_TERMS.copy()
			self.save()

		log.debug(f"Search history initialized. File: {self._history_file}, terms: {self._terms}")

	def _migrate_from_legacy_root_file(self):
		"""Migrates from the oldest layout: userConfig\\logViewer.json"""
		old_file = os.path.join(globalVars.appArgs.configPath, "logViewer.json")
		if not os.path.exists(old_file):
			return
		if os.path.exists(self._history_file):
			return
		log.info(f"Found legacy history file at {old_file}, attempting migration...")
		try:
			shutil.move(old_file, self._history_file)
			log.info(f"Successfully moved legacy file to new location: {self._history_file}")
		except OSError as e:
			log.error(f"Failed to move legacy file using shutil.move: {e}")
			try:
				with open(old_file, 'r', encoding='utf-8') as f:
					terms = json.load(f)
				if isinstance(terms, list) and all(isinstance(t, str) for t in terms):
					self._terms = terms
					self.save()
					os.remove(old_file)
					log.info("Migrated legacy search history by copying and deleting old file.")
				else:
					log.error("Legacy search history file contains invalid data, ignoring.")
			except (OSError, json.JSONDecodeError) as e2:
				log.error(f"Failed fallback migration from legacy file: {e2}")

	def _migrate_from_previous_chai_dir(self):
		"""Migrates from the previous layout: userConfig\\ChaiChaimee\\logViewer.json
		(a plain file, before the ChaiChaimee\\logViewer\\ subfolder existed)."""
		previous_file = os.path.join(globalVars.appArgs.configPath, "ChaiChaimee", "logViewer.json")
		if not os.path.exists(previous_file):
			return
		if os.path.exists(self._history_file):
			return
		log.info(f"Found previous-layout history file at {previous_file}, migrating to {self._history_file}...")
		try:
			shutil.move(previous_file, self._history_file)
			log.info("Successfully migrated previous-layout history file to new subfolder.")
		except OSError as e:
			log.error(f"Failed to move previous-layout file using shutil.move: {e}")
			try:
				with open(previous_file, 'r', encoding='utf-8') as f:
					terms = json.load(f)
				if isinstance(terms, list) and all(isinstance(t, str) for t in terms):
					self._terms = terms
					self.save()
					os.remove(previous_file)
					log.info("Migrated previous-layout search history by copying and deleting old file.")
				else:
					log.error("Previous-layout search history file contains invalid data, ignoring.")
			except (OSError, json.JSONDecodeError) as e2:
				log.error(f"Failed fallback migration from previous-layout file: {e2}")

	def _migrate_from_config(self):
		try:
			old_data = config.conf["LogViewerPlugin"].get("searchHistory")
			if old_data and old_data != "[]":
				try:
					terms = json.loads(old_data)
					if isinstance(terms, list) and all(isinstance(t, str) for t in terms):
						self._terms = terms
						self.save()
						config.conf["LogViewerPlugin"]["searchHistory"] = "[]"
						config.conf.save()
						log.info("Migrated search history from config to JSON file.")
				except json.JSONDecodeError as e:
					log.error(f"Failed to migrate search history: {e}")
		except KeyError:
			pass

	def load(self):
		try:
			if os.path.exists(self._history_file):
				with open(self._history_file, 'r', encoding='utf-8') as f:
					terms = json.load(f)
				if isinstance(terms, list) and all(isinstance(term, str) for term in terms):
					self._terms = terms
				else:
					log.error("Corrupted search history data, resetting to default.")
					self._terms = DEFAULT_TERMS.copy()
					self.save()
			else:
				self._terms = DEFAULT_TERMS.copy()
				self.save()
		except (OSError, json.JSONDecodeError) as e:
			log.error(f"Error loading search history: {e}, resetting to default.")
			self._terms = DEFAULT_TERMS.copy()
			self.save()

	def save(self):
		try:
			os.makedirs(os.path.dirname(self._history_file), exist_ok=True)
			with open(self._history_file, 'w', encoding='utf-8') as f:
				json.dump(self._terms, f, ensure_ascii=False, indent=2)
			log.debug(f"Saved search history to {self._history_file}")
		except OSError as e:
			log.error(f"Error saving search history: {e}")

	def getItems(self):
		return self._terms

	def getItemByText(self, text):
		return next((term for term in self._terms if term.lower() == text.lower()), None)

	def append(self, term):
		if not term:
			return
		existing = self.getItemByText(term)
		if existing:
			self._terms.remove(existing)
		self._terms.insert(0, term)
		if len(self._terms) > 20:
			self._terms.pop()
		self.save()


class PluginSettings:
	"""Holds user-configurable preferences shown on the LogViewer settings panel.
	Stored as JSON under userConfig\\ChaiChaimee\\logViewer\\settings.json, alongside
	the search history file."""

	_instance = None

	@classmethod
	def get(cls):
		if cls._instance is None:
			cls._instance = cls()
		return cls._instance

	def __init__(self):
		self._settings_file = get_settings_file_path()
		self.announceMode = DEFAULT_ANNOUNCE_MODE
		self.quickSearchTerms = DEFAULT_QUICK_SEARCH_TERMS.copy()
		self.load()

	def load(self):
		try:
			if os.path.exists(self._settings_file):
				with open(self._settings_file, 'r', encoding='utf-8') as f:
					data = json.load(f)
				self.announceMode = data.get("announceMode", DEFAULT_ANNOUNCE_MODE)
				terms = data.get("quickSearchTerms")
				if isinstance(terms, list) and all(isinstance(t, str) for t in terms) and terms:
					self.quickSearchTerms = terms
			else:
				self.save()
		except (OSError, json.JSONDecodeError) as e:
			log.error(f"Error loading LogViewer settings: {e}, using defaults.")
			self.announceMode = DEFAULT_ANNOUNCE_MODE
			self.quickSearchTerms = DEFAULT_QUICK_SEARCH_TERMS.copy()

	def save(self):
		try:
			os.makedirs(os.path.dirname(self._settings_file), exist_ok=True)
			with open(self._settings_file, 'w', encoding='utf-8') as f:
				json.dump(
					{
						"announceMode": self.announceMode,
						"quickSearchTerms": self.quickSearchTerms,
					},
					f, ensure_ascii=False, indent=2
				)
			log.debug(f"Saved LogViewer settings to {self._settings_file}")
		except OSError as e:
			log.error(f"Error saving LogViewer settings: {e}")


def initConfiguration():
	confspec = {
		"searchCaseSensitivity": "boolean(default=False)",
		"searchWrap": "boolean(default=True)",
		"searchType": "string(default='NORMAL')",
		"bookmarkCount": "integer(default=1)",
		"lastBootTimestamp": "float(default=0.0)",
	}
	config.conf.spec["LogViewerPlugin"] = confspec
