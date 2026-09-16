# foresight

**A better outlook.** Outlook on the web, reduced to a calm, keyboard-driven
calendar, as a Chrome extension. Optionally tinted to match your Omarchy theme.

Chrome extensions are one of the few things people in a managed workplace
still get to choose. This one is for anyone whose employer left them Outlook
on the web and nothing else. The goal is fewer barriers: one install, one
"install as app", and a calendar that stays out of your way.

Outlook's web calendar is the only client some Microsoft 365 tenants leave
available: if yours blocks third-party OAuth consent, every native calendar app
(Morgen, Notion Calendar, Evolution/GNOME Online Accounts, DavMail) hits an
admin-approval wall, and Graph, EWS, CalDAV and ActiveSync are all either
OAuth-gated or unsupported. That leaves OWA. This makes it bearable.

It strips the suite header, app rail and ribbon, compresses the day grid so a
full working day fits without scrolling, and puts a small action bar top-left
with the controls actually worth keeping.

## Install

### The Chrome extension

1. Install the extension. Until it is on the Chrome Web Store: download or
   clone this repository, open `chrome://extensions`, turn on **Developer
   mode**, click **Load unpacked** and pick the `extension/` folder.
2. Open [outlook.cloud.microsoft/calendar](https://outlook.cloud.microsoft/calendar)
   and sign in.
3. In Chrome's menu choose **Cast, save and share → Install page as app**.

Open it from your launcher like any other app. That window is reduced;
Outlook in an ordinary tab is left exactly as it was.

The extension's toolbar button does the same without installing anything:
it opens the calendar in a window of its own, or focuses that window if it is
already open. Right-click the button for **Options**: hours on screen, start
hour, colour palette, and whether to also apply in normal tabs.

**How it tells the app from a tab.** A window opened as an installed web app,
with `--app=`, or as an extension popup reports `display-mode: standalone`
to the page; a normal tab reports `browser`. Measured on Chromium 152, at
document start and after load. No launcher flag is needed. Anything else on
Outlook is untouched unless you opt in, either with the option or by adding
`?omarchy=1` to a tab's URL.

Works in Chromium-family browsers: Chrome, Brave, Edge, Vivaldi, Helium.
Firefox can load the extension but has no "install as app", so it only gets
the opt-in tab path.

### The userscript (alternative)

`foresight.user.js` is the same code wrapped for a userscript manager, and
is what the local update server serves.

1. Install [Violentmonkey](https://violentmonkey.github.io/get-it/). On
   Chromium you must also enable **Allow user scripts** on the extension's
   details page; Chrome's Manifest V3 requires it.
2. Install `foresight.user.js`.
3. Pin the calendar with the flag in the URL:

```bash
omarchy-webapp-install "Calendar" "https://outlook.cloud.microsoft/calendar/view/workweek?omarchy=1" "calendar"
```

A userscript cannot ask the browser which kind of window it is in, so the
`?omarchy=1` **is required** there. It is read at document-start and stashed
in `sessionStorage`, which is scoped to one tab: the app window keeps it
across every in-app navigation, and an Outlook tab you open normally never
has it. Note that `omarchy-launch-webapp` falls back to Chromium unless your
default browser is Chromium-family, so install the script in whichever
browser actually opens.

### Colours

The extension ships Omarchy's palettes built in; pick one under **Options →
Palette**. On Omarchy itself, choose **follow the desktop theme** instead
and run:

```bash
./setup
```

It installs a user service that serves your current palette on
`127.0.0.1:8787`. The extension polls it and retints OWA to match. Switch
themes and the calendar follows within a few seconds, no relogin, no reload.
Chrome asks once for permission to reach `127.0.0.1` when you pick that option.

Undo with `./setup --uninstall`. The calendar keeps working; it just stops
recolouring.

## Using it

| | |
|---|---|
| `☰` | collapsible sidebar — month picker and per-account calendar lists |
| `New` `D` `W` `M` | new event, and day / week / month views |
| `Alt+Shift+O` | toggle the full Outlook UI back, for anything the bar doesn't cover |
| dot, bottom-right | same toggle, for when you've forgotten the shortcut |

### Keyboard

| key | |
|---|---|
| `c` | quick add — opens the compose ready to type, Enter saves |
| `n` | new event (plain Outlook compose) |
| `t` | today |
| `d` / `w` / `m` | day / week / month view |
| `j` / `k` | next / previous period (vim direction) |
| `s` | show / hide the calendars sidebar |
| `Space` | walk the events on the day you are looking at |
| `Tab` | rotate: bar → slot → nearest event → bar |
| `←` `→` | move along the action bar, once you are on it |

**Tab is a three-stop rotation**, the same three in either direction:

```
   bar  →  selected slot  →  nearest event  →  bar
```

Everything else on the page is out of the tab order, so the next press is
always predictable and the bar is never more than two away.

The arrow keys move the selected slot. Tab from there grabs the event
**nearest where you are** — not the first of the day — and with one focused,
Enter opens it and Delete removes it. To reach a different event, arrow to it
and press Tab: pressing an arrow with an event focused hands focus back to the
grid, which makes that loop work without thinking about it.

`c` is the fast path: it opens Outlook's compose stripped to **title, date,
start, end and Save**, and focuses the title. Type and press Enter — saved.

The date and time fields are ours, and they are there from the moment the box
opens. Tab cycles title → date → start → end → Save and never leaves the box.
Times are forgiving: `1700`, `17`, `17:00`, `5pm` and `5.30pm` all work, and
moving the start drags the end along to keep the duration.

**The box is one width, always.** Outlook's own date picker is never shown:
your values are written into it for the instant of the save, behind a blanked
modal. That is why nothing resizes.

**To create an event somewhere other than now**, either type the date, or
select the slot first. `c`
inherits whatever the grid has selected, so the whole thing stays on the
keyboard:

| | |
|---|---|
| `j` / `k` | move to the next / previous week |
| arrow keys | move the selected slot — up/down by 30 minutes, left/right by a day |
| `c` | compose at that slot |

Then type the title and press Enter. This is Outlook's own selection machinery,
which is why it is reliable — see the note below on what happened when the
script tried to drive the date fields itself instead.

**The date row is editable.** Tab to it and press Enter (or click it) and
Outlook's own picker opens with start date, start time and end time. Tab walks
title → date row → those three → Save, and never leaves the box; typing a value
and tabbing on commits it. Enter from the title or Save files the event.

`n` opens the full compose if you need attendees, recurrence or a body.

`n` opens the same compose untouched, for when you need attendees, recurrence,
a location or a body.

Edit the `SHORTCUTS` map in the script to change them. Keys are ignored while
typing in a field and while a dialog is open, and anything with a modifier is
left alone — so OWA's own `Alt+N`, `Alt+Shift+1..4` and `Ctrl+P` still work.

Click-drag on the grid still creates events at the right time, and invitations
can still be accepted or declined — it's the real Outlook underneath, which is
the entire reason for skinning it rather than rebuilding it.

## Settings

Extension: right-click the toolbar button, **Options**. Settings save as you
change them and sync with your Chrome profile.

| setting | default | |
|---|---|---|
| Hours on screen at once | `14` | how much of the day the grid shows |
| Hour the day opens at | `6` | where the grid sits on load |
| Palette | Outlook's own | a bundled palette, or the Omarchy feed |
| Feed URL | `http://127.0.0.1:8787/theme.json` | where the feed is |
| Also apply in normal browser tabs | off | reduce every Outlook tab, not just the app |
| Start on the calendar | on | an installed Outlook app opens on mail; send it on |

Userscript: Violentmonkey's script editor has a **Values** tab with
`hoursVisible`, `startHour`, `themeUrl` and `themePollMs`. Edited values are
kept forever. Untouched ones follow the code default, so a better default
still reaches an existing install: the script records what it seeded and
only replaces a value that still matches it.

## Layout and build

```
extension/core.js        the reduction, as foresight(env); knows nothing of Chrome or GM
extension/content.js     Chrome shell: storage, app-window detection, theme source
extension/background.js  service worker: feed fetch, toolbar launch, re-injection
extension/options.*      settings page
extension/palettes/      Omarchy's palettes as JSON, plus index.json
userscript/header.js     ==UserScript== block, version stamped by the build
userscript/env.js        Violentmonkey shell: GM storage with seeding, GM XHR
build.sh                 writes foresight.user.js and dist/foresight-<v>.zip
```

The version lives in `extension/manifest.json`; `./build.sh` stamps it into
the userscript and runs `node --check` on everything. Edit the sources, never
`foresight.user.js`.

**Why a content script and not a page script.** The core runs in the
extension's isolated world. Everything it does is DOM: querying, dispatching
events, calling the native value setter on an input. Those cross into the
page fine, and the isolated world is what gives it `chrome.storage`. If a
future Outlook build ever needs page-world access, Manifest V3 allows
`"world": "MAIN"` for a content script; the price is a message bridge for
storage, which is why it is not the default.

**Injection is not guaranteed on a cold start.** Launching an app window with
the browser closed skipped the declared content script in 2 of 4 tries. The
service worker therefore checks every finished Outlook load for the shell's
marker and injects the scripts if it is missing; with that in place, 6 of 6
cold launches were reduced. The fallback lands after load, so those launches
show the full Outlook for a moment first.

## Checking a change to quick add

`verify-quickadd.js` derives the reduction from the bottom up rather than
assuming it. Paste it into the console with a compose open:

```js
Q.candidates()      // every branch holding none of the essential nodes
await Q.reduce()    // hide them one at a time, reverting anything that breaks
await Q.size()      // apply the shipping widths
await Q.verify()    // structure + open the picker + check it fits + close it
await Q.keyboard()  // focusing any field must not displace the others
Q.restore()
```

The checks are structural and never click Save, so running it creates no
events. Run against Outlook's compose on 2026-08-26: **19 candidates, 0
reverted**, and the resulting sizes matched what the script ships — 360px
closed, 560px while the picker is open.

That result is the useful part: the only things the box needs are the title
bar (for close and pop-out), the title input, the date summary and Save.
Everything else Outlook puts in the compose — the whole command bar, the
attendee row, location, Teams, the body editor, Scheduler — can go without
affecting any of them.

Reach for this before changing the strip. Every bug in this feature so far came
from reasoning about the DOM instead of measuring it, and the loop can't make
that mistake: it only believes what it measures.

## How it works, and where it will break

**Chrome removal** keys on `#OwaTitleBar`, `#LeftRail` and
`[data-app-section="Ribbon"]` — ids and data attributes, not the obfuscated
class names, which change between releases.

**The ribbon is parked off-screen, not hidden.** Fluent buttons ignore
programmatic clicks once they have no layout box, so `display: none` would
silently break every button in the action bar.

**Buttons are proxied, not moved.** Re-parenting a node out of OWA's React tree
gets undone on the next render. They're found by `data-unique-id`
(`Ribbon-2532`, `2504`, `2519`, `2505`), which is the Office control number and
identical in every locale — `aria-label` is translated.

**The grid compresses** because OWA positions events as percentages inside a
fixed-height 24-hour column. Shrink the column and every event follows exactly.
The column is found by its inline height rather than a class name.

**Theming** remaps colours by role rather than by token name, because OWA runs
two token systems at once — Fluent v9 (`--colorNeutral*`) and older Fabric slots
(`--neutralPrimary`, `--white`, `--themePrimary`). Greys are projected onto the
palette's neutral ramp by luminance; brand blues take the accent's hue at their
own lightness; everything else is left alone, so status colours and your event
categories survive. A few hundred rules bake colours in as literals rather than
referencing a token, so those get re-emitted as overrides against the same
selectors.

**Light themes won't make OWA light.** The mapping is monotonic in luminance,
which keeps every contrast relationship OWA already had. Inverting it does make
the app follow the theme's mode, but wrecks legibility against event category
colours — which are deliberately left untouched. Set OWA's own Appearance to
match your theme's mode instead. `INVERT_ON_MODE_MISMATCH` is there if you want
to try it anyway.

**Quick add drives the real compose.** Outlook's date row is a
`<div role="button">`, not a field, so it can't be typed into — but clicking it
expands into real inputs, which makes the whole form drivable in-page. Setting
those needs the prototype's native value setter plus `input`/`change` events,
because React ignores direct assignment. OWA always prefills *today* regardless
of the week on screen, so the date is only overridden when today isn't among
the visible `data-column-date` columns.

**The picker is reduced too, and sized to its content.** Outlook packs a
time-zone button, All day / Recurring and a five-row Time suggestions list into
the same popover as the three fields, then sizes the popover to match its
anchor — an inline `--fui-match-target-size` copied from the date row. In the
full compose that row is ~513px and it all fits. In the reduced box the anchor
is 264px while the content is 741, so the popover scrolled sideways: tabbing to
End time scrolled Start date out of view. Dropping everything that is not a
field brings it to 519, and `width: max-content` overrides the anchor match.
The walk stops at the fields' common ancestor, or it takes their labels with it.

Widening it is not enough on its own — the form's own scroll container is 360px
with `overflow-x: hidden` and clips the popover straight back down, so whatever
clips it gets marked `data-owa-unclip`. And once it is wide enough it opens
*below* the rows rather than above them, landing on Save, which is what the
extra `padding-bottom` while the picker is open is for.

**Tab stops are whitelisted, not filtered.** Outlook leaves around fifty
focusable controls in the compose, and focusing one inside the collapsed command
bar visibly grows the box. Everything except the title and Save is given
`tabindex="-1"`. Save sits *earlier* in the DOM than the form, so tabbing
forward off the last stop would leave the modal entirely — Tab is handled
explicitly and cycles within the box instead of relying on DOM order.

**The picker was always opening — twice over, we were hiding it.** This looked
for a long time like a callout that refused to open, and it was nothing of the
sort. Two separate faults, both ours:

1. **The strip hid it.** `stripCompose` marks every sibling branch that leads to
   neither the title nor the date row. The picker mounts *inside* the modal, in
   a branch that pass had already hidden, so it was `display: none` from the
   moment it appeared. Measured: fields present in the DOM with the right
   values, `0 × 0`, hidden by us. The fix has two halves — the picker counts as
   a keeper, and the marking pass *clears* the attribute as well as setting it,
   since a branch can become worth keeping after it was first judged.

   The half that was missed at first: **`stripCompose` runs once per compose**,
   from `quickAdd`. Neither half of that fix does anything unless the strip runs
   *again* once the picker exists, so `update()` re-runs it when a picker field
   appears that has not been seen yet — one pass per opening, guarded on the
   field's identity.
2. **It opened off screen.** Fluent positions the popover with floating-ui
   against the geometry it measured on open — the *unreduced* form. The
   reduction then pulls the anchor hundreds of pixels up, and the popover stays
   put: measured at `y = 1003` in a 756px viewport, below the modal entirely.
   floating-ui recomputes on resize, so one synthetic `resize` event moves it to
   `y = 404`, inside the box. That is the whole fix.

The lesson worth keeping: **check existence separately from visibility.** Every
probe here filtered on `height > 0`, which reports "not open" for a thing that
is open, correct and merely hidden by the caller. That one conflation cost
several rounds and produced two confident, wrong conclusions.

Its fields are named by Fabric — `DatePicker<n>` and `ComboBox<n>-input` — which
is locale-independent and safer than matching a `YYYY-MM-DD` value shape. They
are comboboxes holding a *draft*: moving focus away commits it (which is what
makes Tab work), and so does Enter, which also closes the popover.

**Save is found by Office toolbar metadata, not its label.** It has no id, but
it is the only button in the compose carrying `priorityid="3"` /
`overfloworderid="-3"`, which are the same in every locale. `aria-label="Save"`
is kept as a fallback, and the last resort is the first visible button in the
command bar.

**Quick add strips the form by marking, not matching.** The rows it removes
have no ids and vary in nesting depth, so it walks up from the title to
`Form_Content` and marks every sibling branch leading to neither the title nor
the date — which drops the calendar picker, attendees, location, Teams toggle,
body editor and preview pane in one pass. The modal itself has no explicit
height (it is sized by flex growth), so the same walk marks the chain to shrink
it. Markers are cleared when the compose closes so `n` still gets the full form.

**Expect selector drift.** Microsoft changes this UI regularly, and the hostname
already moved from `outlook.office.com` to `outlook.cloud.microsoft` once during
development. If chrome reappears, inspect the stray element and add its id to
the hide rule.

## Requirements

- A Chromium-family browser, or a userscript manager
- Python 3.11+ for the theme feed (needs `tomllib`), only on Omarchy
- Node, only to run `./build.sh`

## Name

foresight, all lowercase. A calendar is foresight, and the tagline is the
pun: a better outlook. Outlook is a trademark of Microsoft; this project is
independent and not affiliated with Microsoft.

Internal names were not renamed: the DOM attributes (`data-owa-*`), element
ids (`omarchy-owa-*`), storage keys and the `?omarchy=1` flag are what the
CSS, `verify-quickadd.js` and existing installs key on, and changing them
would buy nothing.

## Licence

MIT
