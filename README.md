# Reply to Sender for Microsoft Outlook

An [NVDA](https://www.nvaccess.org/) screen reader add-on that replies
directly — and only — to the original sender of the message you are reading
in Microsoft Outlook.

* Author: Josh Kennedy
* Version: 0.1
* Compatibility: NVDA 2024.1 through 2026.1
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
4. Erases whatever is in the To field and fills it with only the sender's
   address — the same text that is on the clipboard.
5. Moves focus to the very top of the message body and announces:
   "message body, begin typing."

Then just type your reply and send it as usual.

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

This produces `replyToSenderOutlook-0.1.nvda-addon` in the repository root.

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

## License

GNU General Public License version 2. See [LICENSE](LICENSE).
