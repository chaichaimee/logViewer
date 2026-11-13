# __init__.py
# Copyright (C) 2026 Chai Chaimee
# Licensed under GNU General Public License. See COPYING.txt for details.

import globalVars

if globalVars.appArgs.secure:
	from globalPluginHandler import GlobalPlugin
else:
	from .addonCore import GlobalPlugin
