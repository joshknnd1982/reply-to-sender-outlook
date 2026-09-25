# -*- coding: UTF-8 -*-
# Reply to Sender for Microsoft Outlook
# Copyright (C) 2026 Josh Kennedy <joshknnd1982@gmail.com>
# This file is covered by the GNU General Public License version 2.
# See the file LICENSE for more details.

"""Checks GitHub for updates to Reply to Sender for Microsoft Outlook.

Everything the add-on does in Outlook is in appModules/outlook.py, which NVDA
runs only while Outlook is running. Checking for updates has to work while
Outlook is closed too, so it lives in this global plugin, which does nothing
else: no keyboard hooks and nothing in Outlook.
"""

import globalPluginHandler
from scriptHandler import script

import addonHandler

from . import updater

addonHandler.initTranslation()


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	def __init__(self):
		super().__init__()
		# The add-on has no settings of its own, so the update settings get a panel to themselves.
		updater.start(settingsPanel=True)

	def terminate(self):
		updater.stop()
		super().terminate()

	@script(
		# Translators: Description of a command, shown in the Input Gestures dialog.
		description=_("Checks for Reply to Sender for Microsoft Outlook updates"),
		# Translators: name of the input gestures category for this add-on.
		category=_("Reply to Sender for Outlook"),
	)
	def script_checkForUpdates(self, gesture):
		updater.checkForUpdates()
