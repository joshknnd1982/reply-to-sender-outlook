# -*- coding: UTF-8 -*-
# Reply to Sender for Microsoft Outlook
# Copyright (C) 2026 Josh Kennedy <joshknnd1982@gmail.com>
# This file is covered by the GNU General Public License version 2.
# See the file LICENSE for more details.

"""Reply to Sender for Microsoft Outlook.

NVDA+shift+r replies only to the sender of the current Outlook message and
NVDA+shift+e says who that sender is. Both keys are the add-on's only while
classic Outlook has focus; everywhere else NVDA and other add-ons get them.
Message windows Outlook opens are maximized. The work is done in outlookReply.

Up to 0.3 the Outlook part was an app module, appModules/outlook.py. NVDA
runs only one app module for a program, so it replaced Outlook Extended's,
and Outlook Extended's commands, such as Alt+1 for a message's first header
field, stopped working. A global plugin works alongside whichever Outlook app
module NVDA runs, its own or an add-on's.

The plugin also checks GitHub for updates, which has to work while Outlook is
closed too.
"""

import api
import globalPluginHandler
import ui
from logHandler import log
from scriptHandler import script

import addonHandler

from . import outlookReply, updater

addonHandler.initTranslation()

#: NVDA's name for classic Outlook's app module. The new Outlook ("olk") has no COM object model.
OUTLOOK_APP_NAME = "outlook"


def _outlookAppModule(obj):
	"""The app module NVDA runs for classic Outlook when *obj* is in it, or None."""
	try:
		appModule = obj.appModule
		if (appModule.appName or "").lower() == OUTLOOK_APP_NAME:
			return appModule
	except Exception:
		pass
	return None


def _objectModel(appModule):
	"""Outlook's object model, which NVDA's Outlook app module (and any add-on's built on it) gets."""
	try:
		return appModule.nativeOm
	except Exception:
		log.debugWarning("Could not get Outlook's object model", exc_info=True)
		return None


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	def __init__(self):
		super().__init__()
		# The add-on has no settings of its own, so the update settings get a panel to themselves.
		updater.start(settingsPanel=True)

	def terminate(self):
		updater.stop()
		super().terminate()

	def getScript(self, gesture):
		script = super().getScript(gesture)
		if (
			getattr(script, "__func__", None) in _OUTLOOK_SCRIPTS
			and _outlookAppModule(api.getFocusObject()) is None
		):
			# Outside Outlook the key is left to NVDA and other add-ons.
			return None
		return script

	def event_foreground(self, obj, nextHandler):
		# When the user opens a message in its own window (an inspector),
		# make sure that window is maximized before they start working in it.
		appModule = _outlookAppModule(obj)
		if appModule is not None:
			try:
				outlookReply.maximizeIfInspector(_objectModel(appModule), obj)
			except Exception:
				log.debugWarning("Inspector maximize check failed", exc_info=True)
		nextHandler()

	@script(
		# Translators: input help description for the reply to sender script.
		description=_(
			"Replies only to the sender of the current Outlook message: "
			"copies the sender's email address to the clipboard, opens the "
			"reply, replaces the To field with just the sender's address, "
			"and moves focus to the top of the message body."
		),
		gesture="kb:NVDA+shift+r",
		# Translators: name of the input gestures category for this add-on.
		category=_("Reply to Sender for Outlook"),
	)
	def script_replyToSender(self, gesture):
		appModule = _outlookAppModule(api.getFocusObject())
		if appModule is None:
			# Only a key assigned in Input Gestures gets here outside Outlook.
			# Translators: reported when a command of this add-on is used outside classic Outlook.
			ui.message(_("No message to reply to. Open or select a message in classic Outlook."))
			return
		outlookReply.replyToSender(_objectModel(appModule))

	@script(
		# Translators: input help description for the announce sender script.
		description=_(
			"Announces the email address the add-on detects for the sender of "
			"the current Outlook message, without replying to it."
		),
		gesture="kb:NVDA+shift+e",
		# Translators: name of the input gestures category for this add-on.
		category=_("Reply to Sender for Outlook"),
	)
	def script_announceSender(self, gesture):
		appModule = _outlookAppModule(api.getFocusObject())
		if appModule is None:
			# Translators: reported when a command of this add-on is used outside classic Outlook.
			ui.message(_("No message to reply to. Open or select a message in classic Outlook."))
			return
		outlookReply.announceSender(_objectModel(appModule))

	@script(
		# Translators: Description of a command, shown in the Input Gestures dialog.
		description=_("Checks for Reply to Sender for Microsoft Outlook updates"),
		# Translators: name of the input gestures category for this add-on.
		category=_("Reply to Sender for Outlook"),
	)
	def script_checkForUpdates(self, gesture):
		updater.checkForUpdates()


#: The commands that only mean something in Outlook.
_OUTLOOK_SCRIPTS = (GlobalPlugin.script_replyToSender, GlobalPlugin.script_announceSender)
