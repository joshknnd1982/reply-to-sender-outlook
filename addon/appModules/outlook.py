# -*- coding: UTF-8 -*-
# Reply to Sender for Microsoft Outlook
# Copyright (C) 2026 Josh Kennedy <joshknnd1982@gmail.com>
# This file is covered by the GNU General Public License version 2.
# See the file LICENSE for more details.

"""App module extension for Microsoft Outlook.

Adds NVDA+shift+r: reply only to the original sender of the current message.

Because this is an app module, NVDA only runs this code while Microsoft
Outlook is the foreground application, and NVDA unloads it automatically
when Outlook exits (close button, alt+f4, etc.).  No global plugin is used,
so nothing from this add-on remains active outside of Outlook.

The script:
1. Ensures the current Outlook window is maximized.
2. Finds the sender ("From") of the current message and copies the sender's
   SMTP email address to the clipboard.
3. Invokes Outlook's reply command for the message.
4. Waits for the reply window to gain focus (and maximizes it).
5. Replaces whatever is in the To field with only the sender's address
   (the same text that was copied to the clipboard).
6. Moves focus to the very top of the message body and announces
   "message body, begin typing."
"""

import api
import core
import oleacc
import ui
import winUser
from comtypes import COMError
from keyboardHandler import KeyboardInputGesture
from logHandler import log
from NVDAObjects.IAccessible import getNVDAObjectFromEvent
from scriptHandler import script
from windowUtils import findDescendantWindow

import addonHandler

# Subclass NVDA's built-in Outlook app module so all of NVDA's standard
# Outlook support keeps working alongside this add-on.
from nvdaBuiltin.appModules.outlook import AppModule as BuiltinOutlookAppModule

addonHandler.initTranslation()

#: Outlook object model constant: a mail item (OlObjectClass.olMail).
olMail = 43
#: user32 ShowWindow flag to maximize a window.
SW_MAXIMIZE = 3
#: Window class used by both Outlook explorer and inspector top level windows.
OUTLOOK_WINDOW_CLASS = "rctrl_renwnd32"
#: Window class of the Word editor that hosts the message body when composing.
BODY_WINDOW_CLASS = "_WwG"
#: MAPI property tag for the sender's SMTP address (PR_SENDER_SMTP_ADDRESS).
PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
#: How often (ms) and how many times to poll for the reply window.
REPLY_WINDOW_POLL_INTERVAL = 100
REPLY_WINDOW_POLL_MAX_ATTEMPTS = 100  # 10 seconds
BODY_POLL_MAX_ATTEMPTS = 20  # 2 seconds


def _maximizeWindow(hwnd):
	if hwnd and not winUser.user32.IsZoomed(hwnd):
		winUser.user32.ShowWindow(hwnd, SW_MAXIMIZE)


class AppModule(BuiltinOutlookAppModule):
	def event_foreground(self, obj, nextHandler):
		# When the user opens a message in its own window (an inspector),
		# make sure that window is maximized before they start working in it.
		try:
			self._maximizeIfInspector(obj)
		except Exception:
			log.debugWarning("Inspector maximize check failed", exc_info=True)
		superEvent = getattr(super(), "event_foreground", None)
		if superEvent:
			superEvent(obj, nextHandler)
		else:
			nextHandler()

	def _maximizeIfInspector(self, obj):
		hwnd = obj.windowHandle
		if not hwnd or winUser.getClassName(hwnd) != OUTLOOK_WINDOW_CLASS:
			return
		om = self.nativeOm
		if not om:
			return
		try:
			inspector = om.ActiveInspector()
			caption = inspector.Caption if inspector else None
		except COMError:
			return
		# Explorer (main) and inspector (message) windows share a window
		# class, so match the window title against the inspector caption.
		if caption and winUser.getWindowText(hwnd) == caption:
			_maximizeWindow(hwnd)

	def _getCurrentMailItem(self, om):
		"""Return the mail item the user is reading: the item open in the
		active inspector, or the item selected in the active explorer.
		"""
		item = None
		try:
			inspector = om.ActiveInspector()
			if inspector is not None:
				item = inspector.CurrentItem
		except COMError:
			item = None
		if item is None:
			try:
				explorer = om.ActiveExplorer()
				if explorer is not None and explorer.Selection.Count >= 1:
					item = explorer.Selection.Item(1)
			except COMError:
				item = None
		if item is None:
			return None
		try:
			if item.Class != olMail:
				return None
		except COMError:
			return None
		return item

	def _getSenderSmtpAddress(self, item):
		"""Fetch the sender's email address from the message's From field,
		resolving Exchange (X.500) senders to their real SMTP address.
		"""
		try:
			senderType = (item.SenderEmailType or "").upper()
		except COMError:
			senderType = ""
		if senderType == "EX":
			try:
				sender = item.Sender
				if sender is not None:
					exchangeUser = sender.GetExchangeUser()
					if exchangeUser is not None:
						smtp = exchangeUser.PrimarySmtpAddress
						if smtp:
							return smtp
			except COMError:
				pass
			try:
				smtp = item.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
				if smtp:
					return smtp
			except COMError:
				pass
		try:
			return item.SenderEmailAddress or None
		except COMError:
			return None

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
		om = self.nativeOm
		if not om:
			# Translators: reported when the add-on cannot reach Outlook.
			ui.message(_("Unable to connect to Outlook."))
			return
		item = self._getCurrentMailItem(om)
		if item is None:
			# Translators: reported when no message is open or selected.
			ui.message(_("No message to reply to."))
			return
		# Step 1: make sure the current Outlook window is maximized.
		foreground = winUser.getForegroundWindow()
		if winUser.getClassName(foreground) == OUTLOOK_WINDOW_CLASS:
			_maximizeWindow(foreground)
		# Step 2: find the sender's address from the From field and copy it.
		address = self._getSenderSmtpAddress(item)
		if not address:
			# Translators: reported when the sender's address is unavailable.
			ui.message(_("Could not determine the sender's email address."))
			return
		try:
			api.copyToClip(address)
		except Exception:
			log.debugWarning("Could not copy sender address to clipboard", exc_info=True)
		# Step 3: invoke Outlook's reply command for this message.
		try:
			reply = item.Reply()
		except COMError:
			# Translators: reported when the message cannot be replied to.
			ui.message(_("Unable to reply to this message."))
			return
		# Step 4: erase whatever is in the To field and replace it with only
		# the original sender's address (the text now on the clipboard).
		try:
			reply.To = address
			reply.Recipients.ResolveAll()
		except COMError:
			log.debugWarning("Could not set reply recipient", exc_info=True)
		try:
			reply.Display(False)
		except COMError:
			# Translators: reported when the reply window fails to open.
			ui.message(_("Unable to open the reply window."))
			return
		# Step 5: wait for the reply window to gain focus, then move focus
		# into the message body.
		self._waitForReplyWindow(foreground, 0)

	def _waitForReplyWindow(self, previousForeground, attempt):
		hwnd = winUser.getForegroundWindow()
		if (
			hwnd
			and hwnd != previousForeground
			and winUser.getClassName(hwnd) == OUTLOOK_WINDOW_CLASS
		):
			_maximizeWindow(hwnd)
			core.callLater(50, self._focusMessageBody, hwnd, 0)
			return
		if attempt >= REPLY_WINDOW_POLL_MAX_ATTEMPTS:
			# Translators: reported when the reply window never gained focus.
			ui.message(_("The reply window did not appear."))
			return
		core.callLater(
			REPLY_WINDOW_POLL_INTERVAL, self._waitForReplyWindow, previousForeground, attempt + 1
		)

	def _focusMessageBody(self, replyHwnd, attempt):
		try:
			bodyHwnd = findDescendantWindow(
				replyHwnd, visible=True, className=BODY_WINDOW_CLASS
			)
		except LookupError:
			bodyHwnd = None
		if not bodyHwnd:
			if attempt < BODY_POLL_MAX_ATTEMPTS:
				core.callLater(100, self._focusMessageBody, replyHwnd, attempt + 1)
			else:
				# Translators: reported when the message body cannot be found.
				ui.message(_("Could not locate the message body."))
			return
		obj = getNVDAObjectFromEvent(bodyHwnd, oleacc.OBJID_CLIENT, 0)
		if obj:
			obj.setFocus()
		else:
			winUser.setForegroundWindow(replyHwnd)
		core.callLater(300, self._finishInBody)

	def _finishInBody(self):
		# Put the caret at the very top of the message body.
		KeyboardInputGesture.fromName("control+home").send()
		# Translators: announced once focus lands in the reply's message body.
		core.callLater(200, ui.message, _("message body, begin typing."))
