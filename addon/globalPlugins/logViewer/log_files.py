# log_files.py

import logging
import os
import tempfile
import globalVars
from logHandler import log

_lastReportedSource = None


def _reportSource(sourceName, logPath):
	global _lastReportedSource
	if _lastReportedSource == sourceName:
		return
	_lastReportedSource = sourceName
	log.info(f"Log file resolved through {sourceName}: {logPath}")


def getCurrentLogPath():
	for handler in logging.getLogger().handlers:
		if isinstance(handler, logging.FileHandler):
			_reportSource("logger file handler", handler.baseFilename)
			return handler.baseFilename
	configuredPath = getattr(globalVars.appArgs, "logFileName", None)
	if configuredPath:
		_reportSource("appArgs.logFileName", configuredPath)
		return configuredPath
	fallbackPath = os.path.join(tempfile.gettempdir(), "nvda.log")
	_reportSource("temp folder guess", fallbackPath)
	return fallbackPath


def getOldLogPath():
	return os.path.join(os.path.dirname(getCurrentLogPath()), "nvda-old.log")


def readLogText(logPath):
	try:
		with open(logPath, "r", encoding="utf-8", errors="replace") as logFile:
			return logFile.read()
	except OSError as e:
		log.error(f"Could not read log file {logPath}: {e}")
		return None
