# settings_panel.py
# Copyright (C) 2026 Chai Chaimee
# Licensed under GNU General Public License. See COPYING.txt for details.

import wx
import config
from gui import guiHelper
from gui.settingsDialogs import SettingsPanel
import addonHandler

from .config_manager import PluginSettings, ANNOUNCE_MODE_CHOICES

addonHandler.initTranslation()


class LogViewerSettingsPanel(SettingsPanel):
	title = _("Log Viewer")

	def makeSettings(self, settingsSizer):
		settings = PluginSettings.get()
		helper = guiHelper.BoxSizerHelper(self, sizer=settingsSizer)

		self.caseSensitiveCheck = helper.addItem(
			wx.CheckBox(self, label=_("&Case sensitive by default"))
		)
		self.caseSensitiveCheck.SetValue(config.conf["LogViewerPlugin"]["searchCaseSensitivity"])

		self.wrapCheck = helper.addItem(
			wx.CheckBox(self, label=_("&Wrap around by default"))
		)
		self.wrapCheck.SetValue(config.conf["LogViewerPlugin"]["searchWrap"])

		self._announceModeKeys = [choice[0] for choice in ANNOUNCE_MODE_CHOICES]
		announceModeLabels = [choice[1] for choice in ANNOUNCE_MODE_CHOICES]
		self.announceModeCombo = helper.addLabeledControl(
			_("Search result &announcement style:"), wx.Choice, choices=announceModeLabels
		)
		try:
			currentIndex = self._announceModeKeys.index(settings.announceMode)
		except ValueError:
			currentIndex = 0
		self.announceModeCombo.SetSelection(currentIndex)

		self.quickTermsBox = helper.addLabeledControl(
			_("Default &quick-search terms (used by F3), one per line:"),
			wx.TextCtrl,
			style=wx.TE_MULTILINE,
			size=self.scaleSize((300, 100)) if hasattr(self, "scaleSize") else (300, 100)
		)
		self.quickTermsBox.SetValue("\n".join(settings.quickSearchTerms))

	def onSave(self):
		config.conf["LogViewerPlugin"]["searchCaseSensitivity"] = self.caseSensitiveCheck.GetValue()
		config.conf["LogViewerPlugin"]["searchWrap"] = self.wrapCheck.GetValue()

		settings = PluginSettings.get()
		selectedIndex = self.announceModeCombo.GetSelection()
		if 0 <= selectedIndex < len(self._announceModeKeys):
			settings.announceMode = self._announceModeKeys[selectedIndex]

		rawTerms = self.quickTermsBox.GetValue().splitlines()
		cleanedTerms = [term.strip() for term in rawTerms if term.strip()]
		if cleanedTerms:
			settings.quickSearchTerms = cleanedTerms
		settings.save()
