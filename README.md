# Reply to Sender for Microsoft Outlook

An [NVDA](https://www.nvaccess.org/) screen reader add-on that replies
directly — and only — to the original sender of the message you are reading
in Microsoft Outlook.

* Author: Josh Kennedy
* Version: 0.2
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
global keyboard hooks and does nothing outside Outlook.

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

This produces `replyToSenderOutlook-0.2.nvda-addon` in the repository root.

## Repository layout

```
addon/
  manifest.ini          Add-on metadata (name, version, NVDA compatibility)
  appModules/
    outlook.py          The Outlook app module with the NVDA+shift+r script
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
