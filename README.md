# Reply to Sender for Microsoft Outlook

An [NVDA](https://www.nvaccess.org/) screen reader add-on that replies
directly — and only — to the original sender of the message you are reading
in Microsoft Outlook.

* Author: Josh Kennedy
* Version: 0.3
* Compatibility: NVDA 2024.1 through 2026.1
* Requires classic Outlook (Outlook 2024, 2021, 2019 or Microsoft 365
  desktop). The new Outlook for Windows is not supported — it exposes no COM
  automation interface.
* Download: grab the `.nvda-addon` file from the
  [releases page](https://github.com/joshknnd1982/reply-to-sender-outlook/releases)

## Why

Replying in Outlook can leave the To field holding a display name, an
internal Exchange address, or extra recipients. This add-on guarantees your
reply goes to exactly one place: the real email address of the person who
sent the message.

## The command

Press **NVDA+shift+r** while reading a message (opened in its own window or
selected in the message list). The NVDA modifier can be either insert or
caps lock — both work. NVDA then:

1. Ensures the current Outlook window is maximized.
2. Finds the message's From field and copies the sender's email address to
   the clipboard. For Exchange / Microsoft 365 senders, the real SMTP
   address is resolved instead of the internal Exchange address.
3. Invokes Outlook's reply command and waits for the reply window to gain
   focus, maximizing it as well.
4. Erases whatever is in the To field — and clears CC and BCC — then fills
   To with only the sender's address, the same text that is on the
   clipboard.
5. Moves focus to the very top of the message body and announces:
   "message body, begin typing."

Then just type your reply and send it as usual.

Press **NVDA+shift+e** to hear the address the add-on detects for the
current message's sender without replying, which is handy for confirming it
reads the From field correctly.

## How the sender is found

The From field is read through Outlook's COM object model, trying several
sources in order of reliability and accepting only a genuine email address:

1. `PR_SENT_REPRESENTING_SMTP_ADDRESS` — what the From field actually shows,
   including mail sent on behalf of somebody else.
2. `PR_SENDER_SMTP_ADDRESS` — the account that sent the message.
3. The `Sender` / `SentOnBehalfOf` address entries, resolved via
   `GetExchangeUser`, `GetExchangeDistributionList`, the entry's
   `PR_SMTP_ADDRESS`, or a matching contact.
4. `SenderEmailAddress`, used only when it holds a real SMTP address.
5. As a last resort, the recipient Outlook itself resolved for the reply.

An Exchange X.500 name such as `/O=EXCHANGELABS/OU=.../CN=abc123` is never
accepted, since it is not a usable address.

## Scope

The add-on is implemented as an Outlook app module, so NVDA only runs it
while Microsoft Outlook is the focused application. When Outlook closes
(close button, alt+f4), NVDA unloads the add-on automatically. It adds no
global keyboard hooks and does nothing else outside Outlook: its one global
plugin only checks for updates, so that it can do so while Outlook is closed.

## Updates

The add-on checks for updates. Once a day, a little after NVDA starts, the add-on asks its GitHub repository, [github.com/joshknnd1982/reply-to-sender-outlook](https://github.com/joshknnd1982/reply-to-sender-outlook), whether a newer version has been released, and says nothing unless there is one. When there is, a dialog shows what's new in a box you can read line by line, and offers to download and install it. The download must match the release's SHA-256 checksum. Then NVDA asks you to confirm the installation and offers to restart. Your settings are kept.

To check yourself, open the NVDA menu, choose **Tools**, then **Check for add-on updates**, and choose **Reply to Sender for Microsoft Outlook...**. Or press **Check for updates now** in the add-on's settings: NVDA menu, Preferences, Settings, **Reply to Sender for Microsoft Outlook**. You can also assign a gesture to **Checks for Reply to Sender for Microsoft Outlook updates** in NVDA's Input Gestures dialog, under **Reply to Sender for Outlook**. To stop the daily check, clear **Check for Reply to Sender for Microsoft Outlook updates automatically** in the same settings panel.

## Installation

1. Download the latest `replyToSenderOutlook-x.y.nvda-addon` file from the
   [releases page](https://github.com/joshknnd1982/reply-to-sender-outlook/releases).
2. Press enter on the downloaded file and confirm the installation in NVDA.
3. Restart NVDA when prompted.

## Building from source

Requires Python 3. From the repository root:

```bash
python build.py
```

This produces `replyToSenderOutlook-0.3.nvda-addon` and its `.sha256` checksum
file in the repository root. Upload both to the GitHub release: the update check
reads the release's tag, such as `v0.3`, and checks the download against the
checksum.

## Repository layout

```
addon/
  manifest.ini          Add-on metadata (name, version, NVDA compatibility)
  appModules/
    outlook.py          The Outlook app module with the NVDA+shift+r script
  globalPlugins/
    replyToSenderOutlook/
      __init__.py       Starts the update check, which must work outside Outlook
      updater.py        The GitHub update check, shared by all of joshknnd1982's
                        add-ons; keep it identical
  doc/
    en/
      readme.html       User documentation bundled with the add-on
build.py                Builds the .nvda-addon package
```

## Notes for developers

The app module subclasses NVDA's built-in Outlook app module
(`nvdaBuiltin.appModules.outlook`), so all of NVDA's standard Outlook
support keeps working. The sender address and the reply are obtained through
Outlook's COM object model (`nativeOm`), which is more reliable than
scraping the visible From field; the reply window and message body are then
tracked through window events so focus lands at the top of the body.

NVDA exposes `nativeOm` as a **dynamic** COM dispatch object. On a dynamic
dispatch, a property that is missing or that fails raises `AttributeError`
or `TypeError` rather than `COMError`, so every Outlook call is guarded with
a broad `except Exception`. Catching only `COMError` lets those errors
escape and abort the script — that was the cause of the sender lookup
failing in 0.1.

## Changelog

### 0.3

* Checks GitHub for updates once a day and offers to install them. To check
  yourself, choose Tools, Check for add-on updates, Reply to Sender for
  Microsoft Outlook in the NVDA menu, or press Check for updates now in the
  add-on's new settings panel.

### 0.2

* Fixed the sender not being detected. Outlook calls were guarded against
  `COMError` only, but NVDA's dynamic COM dispatch raises `AttributeError` /
  `TypeError`, which aborted the whole script before an address was found.
* The sender lookup now tries the sent-representing and sender MAPI
  properties, the `Sender` and `SentOnBehalfOf` address entries, the string
  properties, and finally Outlook's own resolved reply recipient.
* Exchange X.500 names are no longer accepted as email addresses; only a
  genuine address is used.
* The message the user is reading is located via `ActiveWindow` first, so an
  unrelated message window left open in the background is no longer replied
  to by mistake.
* CC and BCC are now cleared along with the To field.
* Added NVDA+shift+e to announce the detected sender address.

### 0.1

* First release.

## License

GNU General Public License version 2. See [LICENSE](LICENSE).
