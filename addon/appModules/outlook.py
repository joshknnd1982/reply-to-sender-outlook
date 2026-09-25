# -*- coding: UTF-8 -*-
# Reply to Sender for Microsoft Outlook
# Copyright (C) 2026 Josh Kennedy <joshknnd1982@gmail.com>
# This file is covered by the GNU General Public License version 2.
# See the file LICENSE for more details.

"""App module extension for Microsoft Outlook.

Adds NVDA+shift+r: reply only to the original sender of the current message.

Because this is an app module, NVDA only runs this code while Microsoft
Outlook is the foreground application, and NVDA unloads it automatically
when Outlook exits (close button, alt+f4, etc.).  The add-on's only global
plugin, globalPlugins/replyToSenderOutlook, checks GitHub for updates and does
nothing else, so none of the Outlook behavior below is active outside Outlook.

The script:
1. Ensures the current Outlook window is maximized.
2. Finds the sender ("From") of the current message and copies the sender's
   SMTP email address to the clipboard.
3. Invokes Outlook's reply command for the message.
4. Waits for the reply window to gain focus (and maximizes it).
5. Replaces whatever is in the To field with only the sender's address,
   and clears CC and BCC.
6. Moves focus to the very top of the message body and announces
   "message body, begin typing."

Note on error handling: NVDA exposes Outlook through a *dynamic* COM
dispatch object (nativeOm).  On a dynamic dispatch, a property that is
missing or that fails raises AttributeError or TypeError rather than
COMError, so every Outlook call below is guarded with a broad ``except
Exception``.  Catching only COMError lets those errors escape and abort the
whole script, which is what previously stopped the sender being found.
"""

import api
import core
import oleacc
import ui
import winUser
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

#: user32 ShowWindow flag to maximize a window.
SW_MAXIMIZE = 3
#: Window class used by both Outlook explorer and inspector top level windows.
OUTLOOK_WINDOW_CLASS = "rctrl_renwnd32"
#: Window class of the Word editor that hosts the message body when composing.
BODY_WINDOW_CLASS = "_WwG"

# MAPI property tags, used through PropertyAccessor.
#: PR_SENT_REPRESENTING_SMTP_ADDRESS - the address behind the visible "From"
#: field, including mail sent on behalf of somebody else.
PR_SENT_REPRESENTING_SMTP_ADDRESS = (
	"http://schemas.microsoft.com/mapi/proptag/0x5D02001F"
)
#: PR_SENDER_SMTP_ADDRESS - the address of the account that actually sent it.
PR_SENDER_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
#: PR_SMTP_ADDRESS - the SMTP address of an AddressEntry.
PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"

#: OlObjectClass values that can never be replied to, so the script can
#: report a useful message instead of failing later.  Anything not listed
#: here is attempted, so unusual but repliable item types still work.
NON_REPLIABLE_CLASSES = frozenset((
	26,  # olAppointment
	40,  # olContact
	42,  # olJournal
	44,  # olNote
	48,  # olTask
	69,  # olDistributionList
))

#: How often (ms) and how many times to poll for the reply window.
REPLY_WINDOW_POLL_INTERVAL = 100
REPLY_WINDOW_POLL_MAX_ATTEMPTS = 100  # 10 seconds
BODY_POLL_MAX_ATTEMPTS = 20  # 2 seconds


def _maximizeWindow(hwnd):
	if hwnd and not winUser.user32.IsZoomed(hwnd):
		winUser.user32.ShowWindow(hwnd, SW_MAXIMIZE)


def _isUsableAddress(value):
	"""Return True only for something that can actually be sent to.

	Exchange senders often report an X.500 distinguished name such as
	``/O=EXCHANGELABS/OU=.../CN=RECIPIENTS/CN=abc123``.  That is not an
	email address; putting it in the To field produces a recipient Outlook
	cannot resolve, so it must never be accepted as the sender's address.
	"""
	if not value or not isinstance(value, str):
		return False
	value = value.strip()
	return "@" in value and not value.startswith("/")


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
		except Exception:
			return
		# Explorer (main) and inspector (message) windows share a window
		# class, so match the window title against the inspector caption.
		if caption and winUser.getWindowText(hwnd) == caption:
			_maximizeWindow(hwnd)

	def _getCurrentMailItem(self, om):
		"""Return the message the user is actually reading.

		ActiveWindow is asked first because it reports whichever window
		genuinely has focus.  Asking ActiveInspector first (as this add-on
		used to) can return a message window that is still open in the
		background while the user is really reading in the reading pane,
		which replies to the wrong message.
		"""
		item = None
		for getter in (
			lambda: om.ActiveWindow(),
			lambda: om.ActiveInspector(),
			lambda: om.ActiveExplorer(),
		):
			try:
				window = getter()
			except Exception:
				continue
			if window is None:
				continue
			# An inspector exposes CurrentItem; an explorer exposes Selection.
			try:
				item = window.CurrentItem
			except Exception:
				item = None
			if item is None:
				try:
					selection = window.Selection
					if selection is not None and selection.Count >= 1:
						item = selection.Item(1)
				except Exception:
					item = None
			if item is not None:
				break
		if item is None:
			return None
		try:
			itemClass = item.Class
		except Exception:
			itemClass = None
		if itemClass in NON_REPLIABLE_CLASSES:
			return None
		return item

	def _addressFromEntry(self, entry):
		"""Resolve an Outlook AddressEntry to a real SMTP address."""
		if entry is None:
			return None
		# Exchange user (the common case on Exchange and Microsoft 365).
		try:
			exchangeUser = entry.GetExchangeUser()
			if exchangeUser is not None:
				address = exchangeUser.PrimarySmtpAddress
				if _isUsableAddress(address):
					return address
		except Exception:
			pass
		# Exchange distribution list.
		try:
			distList = entry.GetExchangeDistributionList()
			if distList is not None:
				address = distList.PrimarySmtpAddress
				if _isUsableAddress(address):
					return address
		except Exception:
			pass
		# The SMTP address stored directly on the address entry.
		try:
			address = entry.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
			if _isUsableAddress(address):
				return address
		except Exception:
			pass
		# A matching contact in the address book.
		try:
			contact = entry.GetContact()
			if contact is not None:
				address = contact.Email1Address
				if _isUsableAddress(address):
					return address
		except Exception:
			pass
		# Plain SMTP entries expose the address directly.
		try:
			address = entry.Address
			if _isUsableAddress(address):
				return address
		except Exception:
			pass
		return None

	def getSenderAddress(self, item):
		"""Return the email address behind the message's From field.

		Several strategies are tried in order of reliability.  Each one is
		fully guarded, so a failure moves on to the next rather than
		aborting, and only a genuine email address is ever returned.
		"""
		# The MAPI properties are the most reliable source and work for
		# Exchange, Microsoft 365 and internet mail alike.  Sent-representing
		# comes first because that is what the From field actually shows for
		# messages sent on behalf of somebody else.
		for tag in (PR_SENT_REPRESENTING_SMTP_ADDRESS, PR_SENDER_SMTP_ADDRESS):
			try:
				address = item.PropertyAccessor.GetProperty(tag)
			except Exception:
				continue
			if _isUsableAddress(address):
				log.debug(f"Sender address resolved from MAPI property {tag}")
				return address.strip()
		# The Sender and SentOnBehalfOf address entries.
		for name in ("Sender", "SentOnBehalfOf"):
			try:
				entry = getattr(item, name)
			except Exception:
				continue
			address = self._addressFromEntry(entry)
			if address:
				log.debug(f"Sender address resolved from {name}")
				return address.strip()
		# Finally the simple string properties, which are usable only when
		# they hold real SMTP addresses rather than an Exchange X.500 name.
		for name in ("SenderEmailAddress", "SentOnBehalfOfName"):
			try:
				address = getattr(item, name)
			except Exception:
				continue
			if _isUsableAddress(address):
				log.debug(f"Sender address resolved from {name}")
				return address.strip()
		return None

	def _addressFromReply(self, reply):
		"""Last resort: read the recipient Outlook itself worked out.

		Outlook has already resolved who this reply goes to, so its own
		answer is a dependable fallback when the message properties cannot
		be read.
		"""
		try:
			recipients = reply.Recipients
			count = recipients.Count
		except Exception:
			return None
		for index in range(1, count + 1):
			try:
				recipient = recipients.Item(index)
			except Exception:
				continue
			try:
				address = self._addressFromEntry(recipient.AddressEntry)
			except Exception:
				address = None
			if address:
				log.debug("Sender address resolved from the reply's recipients")
				return address.strip()
			try:
				address = recipient.Address
			except Exception:
				address = None
			if _isUsableAddress(address):
				log.debug("Sender address resolved from a reply recipient address")
				return address.strip()
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
		# Step 2: find the sender's address from the From field.
		address = self.getSenderAddress(item)
		# Step 3: invoke Outlook's reply command for this message.
		try:
			reply = item.Reply()
		except Exception:
			log.debugWarning("Outlook refused to create a reply", exc_info=True)
			# Translators: reported when the message cannot be replied to.
			ui.message(_("Unable to reply to this message."))
			return
		if not address:
			# The message properties did not give up an address, so fall back
			# to the recipient Outlook resolved for the reply itself.
			address = self._addressFromReply(reply)
		if not address:
			log.debugWarning("Every sender address strategy failed")
			# Translators: reported when the sender's address is unavailable.
			ui.message(_("Could not determine the sender's email address."))
			return
		# Step 4: copy the sender's address to the clipboard.
		try:
			api.copyToClip(address)
		except Exception:
			log.debugWarning("Could not copy sender address to clipboard", exc_info=True)
		# Step 5: erase everything in the To, CC and BCC fields and put back
		# only the original sender, so the reply can reach nobody else.
		try:
			reply.To = address
			reply.CC = ""
			reply.BCC = ""
			reply.Recipients.ResolveAll()
		except Exception:
			log.debugWarning("Could not set the reply recipient", exc_info=True)
			# Translators: reported when the To field could not be replaced.
			ui.message(_("Could not set the reply recipient."))
			return
		try:
			reply.Display(False)
		except Exception:
			log.debugWarning("Could not display the reply", exc_info=True)
			# Translators: reported when the reply window fails to open.
			ui.message(_("Unable to open the reply window."))
			return
		# Step 6: wait for the reply window to gain focus, then move focus
		# into the message body.
		self._waitForReplyWindow(foreground, 0)

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
		address = self.getSenderAddress(item)
		if address:
			ui.message(address)
		else:
			# Translators: reported when the sender's address is unavailable.
			ui.message(_("Could not determine the sender's email address."))

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
