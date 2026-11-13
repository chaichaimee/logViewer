# search_dialog.py

import wx
import ui
from gui import guiHelper
import addonHandler

from .config_manager import PluginSettings, SearchHistory
from .search_logic import SearchRequest, SearchType, compileSearchPattern, describeSearchFailure

addonHandler.initTranslation()


class LogSearchDialog(wx.Dialog):

	def __init__(self, parent):
		super().__init__(parent, title=_("Search in NVDA Log"))
		self.searchHistory = SearchHistory.get()
		self.settings = PluginSettings.get()
		self.request = None
		self._buildGui()
		self.Bind(wx.EVT_BUTTON, self.onOk, id=wx.ID_OK)
		self.Bind(wx.EVT_BUTTON, self.onCancel, id=wx.ID_CANCEL)

	def _buildGui(self):
		mainSizer = wx.BoxSizer(wx.VERTICAL)
		sHelper = guiHelper.BoxSizerHelper(self, orientation=wx.VERTICAL)

		rememberedTerms = self.searchHistory.getItems()
		self.searchBox = sHelper.addLabeledControl(
			_("Search &for:"), wx.ComboBox, choices=rememberedTerms, style=wx.CB_DROPDOWN
		)
		if rememberedTerms:
			self.searchBox.Select(0)

		self.searchTypeBox = sHelper.addItem(wx.RadioBox(
			self,
			label=_("Search type:"),
			choices=SearchType.getSearchTypes(),
			majorDimension=1,
			style=wx.RA_SPECIFY_ROWS
		))
		self.searchTypeBox.SetSelection(SearchType.getIndexByName(self.settings.searchType))

		self.caseSensitiveCheck = sHelper.addItem(wx.CheckBox(self, label=_("Case &sensitive")))
		self.caseSensitiveCheck.SetValue(self.settings.searchCaseSensitive)

		self.wrapCheck = sHelper.addItem(wx.CheckBox(self, label=_("&Wrap around")))
		self.wrapCheck.SetValue(self.settings.searchWrap)

		buttons = self.CreateButtonSizer(wx.OK | wx.CANCEL)
		self.FindWindowById(wx.ID_OK).SetLabel(_("&Find"))
		sHelper.addDialogDismissButtons(buttons)

		mainSizer.Add(sHelper.sizer, border=guiHelper.BORDER_FOR_DIALOGS, flag=wx.ALL)
		mainSizer.Fit(self)
		self.SetSizer(mainSizer)
		self.CentreOnScreen()
		self.searchBox.SetFocus()

	def onOk(self, evt):
		term = self.searchBox.GetValue().strip()
		if not term:
			ui.message(_("Search term cannot be empty"))
			self.searchBox.SetFocus()
			return
		searchType = SearchType.getByIndex(self.searchTypeBox.GetSelection())
		caseSensitive = self.caseSensitiveCheck.GetValue()
		compiledPattern, patternStatus = compileSearchPattern(term, caseSensitive, searchType)
		if compiledPattern is None:
			ui.message(describeSearchFailure(patternStatus))
			self.searchBox.SetFocus()
			return
		wrap = self.wrapCheck.GetValue()
		self.searchHistory.append(term)
		self.settings.setSearchOptions(caseSensitive, wrap, searchType.name)
		self.request = SearchRequest(term, caseSensitive, wrap, searchType)
		self.EndModal(wx.ID_OK)

	def onCancel(self, evt):
		self.EndModal(wx.ID_CANCEL)
