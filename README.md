# owa-minimal

Outlook on the web, reduced to a calendar — and tinted to match your Omarchy theme.

Outlook's web calendar is the only client some Microsoft 365 tenants leave
available: if yours blocks third-party OAuth consent, every native calendar app
(Morgen, Notion Calendar, Evolution/GNOME Online Accounts, DavMail) hits an
admin-approval wall, and Graph, EWS, CalDAV and ActiveSync are all either
OAuth-gated or unsupported. That leaves OWA. This makes it bearable.

It strips the suite header, app rail and ribbon, compresses the day grid so a
full working day fits without scrolling, and puts a small action bar top-left
with the controls actually worth keeping.

## Install

### The userscript (all you need)

1. Install a userscript manager — [Violentmonkey](https://violentmonkey.github.io/get-it/)
   works on both Chromium and Firefox/Zen.
   On Chromium you must also enable **Allow user scripts** on the extension's
   details page (or Developer Mode); Chrome's Manifest V3 requires it.
2. Install `owa-minimal.user.js`.
3. Open your calendar.

That's the whole thing. Everything below is optional.

### The theme feed (Omarchy only)

```bash
./setup
```

Installs a user service that serves your current palette on
`127.0.0.1:8787`. The userscript polls it and retints OWA to match. Switch
themes and the calendar follows within a few seconds — no relogin, no reload.

Undo with `./setup --uninstall`. The userscript keeps working; it just stops
recolouring.

### Pinning it as an app

```bash
omarchy-webapp-install "Calendar" "https://outlook.cloud.microsoft/calendar/view/workweek?omarchy=1" "calendar"
```

Note that `omarchy-launch-webapp` falls back to Chromium unless your default
browser is Chromium-family — so install the userscript in whichever browser
actually opens.

**The `?omarchy=1` is required.** A userscript matches on URL and can't tell
which window it's in, so without a marker it would reshape every Outlook tab in
the browser, not just the app. The flag is read at document-start and stashed in
`sessionStorage`, which is scoped to one tab: the app window keeps it across
every in-app navigation, and an Outlook tab you open normally never has it.

Add the flag by hand to any URL to opt a regular tab in.

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

`c` is the fast path: it opens Outlook's compose stripped to **title, date,
start and end**, with only Save left, and focuses the title. Tab moves title →
date → start → end; Enter saves from anywhere in the box. Shift+Enter stays a
newline.

`n` opens the same compose untouched, for when you need attendees, recurrence,
a location or a body.

Edit the `SHORTCUTS` map in the script to change them. Keys are ignored while
typing in a field and while a dialog is open, and anything with a modifier is
left alone — so OWA's own `Alt+N`, `Alt+Shift+1..4` and `Ctrl+P` still work.

Click-drag on the grid still creates events at the right time, and invitations
can still be accepted or declined — it's the real Outlook underneath, which is
the entire reason for skinning it rather than rebuilding it.

## Settings

Violentmonkey's script editor has a **Values** tab with:

| key | default | |
|---|---|---|
| `hoursVisible` | `14` | hours on screen at once |
| `startHour` | `6` | where the grid sits on load |
| `themeUrl` | `http://127.0.0.1:8787/theme.json` | palette feed |
| `themePollMs` | `3000` | how often to check it |

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

**Quick add uses its own date and time fields.** Outlook's real ones live in a
callout that closes on any outside interaction, so they cannot be tabbed
through; ours stand in and are written across on save. Three things that are
easy to get wrong there: the callout only opens on a full
`pointerdown → mousedown → mouseup → click` (a bare `.click()` does nothing);
Outlook's fields are comboboxes that commit a draft value only on Enter *while
focused*, so `reactSet` alone silently saves the old time; and when searching
for those fields, our own row has to be excluded, because a native
`input[type=date]` reports the same `YYYY-MM-DD` shape and wins on DOM order.

**Quick add strips the form by marking, not matching.** The rows it removes
have no ids and vary in nesting depth, so it walks up from the title to
`Form_Content` and marks every sibling branch leading to neither the title nor
the date — which drops the calendar picker, attendees, location, Teams toggle,
body editor and preview pane in one pass. The modal itself has no explicit
height (it is sized by flex growth), so the same walk marks the chain to shrink
it. Markers are cleared when the compose closes so `n` still gets the full form.

**The Save button is the weakest selector in the project** — it has no id, only
`aria-label="Save"`, which is translated. Quick add will not save on a
non-English Outlook until that gains a fallback.

**Expect selector drift.** Microsoft changes this UI regularly, and the hostname
already moved from `outlook.office.com` to `outlook.cloud.microsoft` once during
development. If chrome reappears, inspect the stray element and add its id to
the hide rule.

## Requirements

- A userscript manager
- Python 3.11+ for the theme feed (needs `tomllib`) — not needed for the
  userscript alone
- Omarchy, for the theme feed only

## Licence

MIT
