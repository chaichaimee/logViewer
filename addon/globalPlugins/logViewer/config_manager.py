# config_manager.py
# Copyright (C) 2026 Chai Chaimee
# Licensed under GNU General Public License. See COPYING.txt for details.

import json
import os
import shutil
from logHandler import log
import config
import globalVars


def get_history_file_path():
	base_dir = globalVars.appArgs.configPath
	chai_dir = os.path.join(base_dir, "ChaiChaimee")
	if not os.path.exists(chai_dir):
		try:
			os.makedirs(chai_dir)
		except OSError as e:
			log.error(f"Could not create directory {chai_dir}: {e}")
	return os.path.join(chai_dir, "logViewer.json")


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

		if not self._terms:
			self._migrate_from_old_file()

		if not self._terms:
			self._migrate_from_config()

		self.save()
		log.debug(f"Search history initialized. File: {self._history_file}, terms: {self._terms}")

	def _migrate_from_old_file(self):
		old_file = os.path.join(globalVars.appArgs.configPath, "logViewer.json")
		if not os.path.exists(old_file):
			log.debug("No old history file found at %s", old_file)
			return
		log.info(f"Found old history file at {old_file}, attempting migration...")
		try:
			shutil.move(old_file, self._history_file)
			log.info(f"Successfully moved old file to new location: {self._history_file}")
			self.load()
		except Exception as e:
			log.error(f"Failed to move old file using shutil.move: {e}")
			try:
				with open(old_file, 'r', encoding='utf-8') as f:
					terms = json.load(f)
				if isinstance(terms, list) and all(isinstance(t, str) for t in terms):
					self._terms = terms
					self.save()
					os.remove(old_file)
					log.info("Migrated search history by copying and deleting old file.")
				else:
					log.error("Old search history file contains invalid data, ignoring.")
			except Exception as e2:
				log.error(f"Failed fallback migration: {e2}")

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
				except Exception as e:
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
					log.error("Corrupted search history data, resetting to empty list.")
					self._terms = []
			else:
				self._terms = []
		except Exception as e:
			log.error(f"Error loading search history: {e}, resetting to empty list.")
			self._terms = []

	def save(self):
		try:
			os.makedirs(os.path.dirname(self._history_file), exist_ok=True)
			with open(self._history_file, 'w', encoding='utf-8') as f:
				json.dump(self._terms, f, ensure_ascii=False, indent=2)
			log.debug(f"Saved search history to {self._history_file}")
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


def initConfiguration():
	confspec = {
		"searchCaseSensitivity": "boolean(default=False)",
		"searchWrap": "boolean(default=True)",
		"searchType": "string(default='NORMAL')",
		"bookmarkCount": "integer(default=1)",
	}
	config.conf.spec["LogViewerPlugin"] = confspec