# Chrome Web Store listing

Everything the developer dashboard asks for, in the order it asks.

## Where

- Register (US$5 once, choose **non-trader**): https://chrome.google.com/webstore/devconsole/register
- Dashboard: https://chrome.google.com/webstore/devconsole
- Upload: **New item** → drop `dist/foresight-<version>.zip`
- Policies worth reading once: https://developer.chrome.com/docs/webstore/program-policies

## Store listing

**Name**: foresight

**Summary** (132 characters max):

> A better outlook: your Outlook Web calendar reduced to a calm, keyboard-driven app. Ordinary Outlook tabs are left alone.

**Description**:

> Outlook on the web, reduced to the calendar — and the calendar reduced to what a calendar is for.
>
> foresight strips the suite header, the app rail and the ribbon, fits a full working day on screen without scrolling, and leaves one bar: the week at the left, the controls worth keeping at the right. Everything is on the keyboard: `c` opens a quick add ready to type, Enter saves; `j` and `k` move through the weeks; the arrow keys move the selected slot; `?` shows the rest.
>
> It works on the Outlook you already have. Nothing to sign into, nothing to sync, no access to your account: it reshapes the page in front of you and leaves your data where it is.
>
> Install Outlook as an app (Chrome menu → Cast, save and share → Install page as app) and foresight takes over that window; your ordinary Outlook tabs are left exactly as they were.
>
> Colours are yours to set: Outlook's own, one of the bundled palettes, or — on Omarchy — the desktop theme, followed live as it changes, with event colours mapped for contrast.
>
> Made for anyone whose workplace left them Outlook on the web and nothing else.

**Category**: Productivity

**Language**: English

## Graphics

- Icon 128×128: `extension/icons/icon128.png` (already in the package)
- Screenshots, 1280×800 or 640×400, 1 to 5: take them from the app window
  1. the week view in your theme — this is the one that sells it
  2. quick add open, a title typed
  3. the `?` keys dialog
  4. the options page
- Small promo tile 440×280: optional; the icon on the palette background with the name is enough

## Privacy tab

**Single purpose**:

> Reduces Outlook on the web to a keyboard-driven calendar and recolours it.

**Permission justifications**:

- `storage` — keeps the user's settings (which view to open, colours, font).
- `scripting` — Chrome skips the declared content script on some cold launches of an installed web app; the service worker re-injects it after the page loads. Measured, not hypothetical.
- Host permissions for `outlook.office.com`, `outlook.office365.com`, `outlook.cloud.microsoft` — the pages the extension reshapes. It runs nowhere else.
- Optional host permissions for `127.0.0.1` and `localhost` — the local theme feed (Omarchy). Requested only when the user enables the feed in the options.

**Remote code**: No, no remote code is used.

**Data usage**: tick nothing — no data is collected. Certify the three statements.

**Privacy policy URL**: https://github.com/pedrosekine/foresight/blob/master/PRIVACY.md

## After submitting

Review takes from a day to a couple of weeks. A new version is the same zip
with a higher `version` in `manifest.json`, uploaded to the existing item.
