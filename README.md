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

`c` is the fast path: it opens Outlook's compose stripped to **title, date/time
and Save**, and focuses the title. Type and press Enter — saved.

Tab cycles title → Save and back, and never leaves the box.

**The date row is shown but not editable from the reduced box.** Outlook's date
picker will not open while the box is reduced — see below. Use `n` for anything
that is not at the slot shown.

Own date and time fields are written and tabbable but disabled behind
`QUICK_ADD_FIELDS`, because writing their values back into Outlook could not be
made reliable — see below.

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

Edited values are kept forever. Untouched ones follow the code default, so a
better default still reaches an existing install — the script records what it
seeded and only replaces a value that still matches it.

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

**Tab stops are whitelisted, not filtered.** Outlook leaves around fifty
focusable controls in the compose, and focusing one inside the collapsed command
bar visibly grows the box. Everything except the title, our three fields and
Save is given `tabindex="-1"`. Note that Save sits *earlier* in the DOM than the
form, so tabbing forward off the last field leaves the modal entirely — the end
field hands focus to Save explicitly.

**The fields are text, not `input[type=date|time]`.** The native ones split into
`hh` / `mm` / `AM-PM` segments that are each their own tab stop, which made
reaching the end time six presses instead of three.

**The reduced box and Outlook's date picker are mutually exclusive.** With the
hidden rows set to `display: none`, the picker never opens — not by script, and
not by a real mouse click either. Unset that one rule and it opens every time.
It does not mount inside a hidden branch (its ancestors carry no marker), so the
mechanism is still unexplained; the boundary is simply reproducible. Toggling
the rule at the moment of opening did not work either.

**Outlook's date fields cannot be driven reliably.** They live in a callout
that only opens when its row is laid out normally and has been for a moment:
hide the row, move it off-screen, fade it or collapse its height and the callout
never opens, and a row that started hidden never opens at all. Even with the row
restored and visible, the same pointer sequence opened it one moment and did
nothing the next. Since a failed write saves silently at the wrong time rather
than erroring, quick add keeps Outlook's own row instead. A real click or key
press on it works every time; only synthetic ones are unreliable.

**Quick add uses its own date and time fields.** Outlook's real ones live in a
callout that closes on any outside interaction, so they cannot be tabbed
through; ours stand in and are written across on save. That callout will not
open unless Outlook's date row is laid out normally and on screen — hidden,
moved off-screen or faded all stop it — so the row is hidden outright and put
back for the instant of the commit, with the modal blanked meanwhile. Three things that are
easy to get wrong there: the callout only opens on a full
`pointerdown → mousedown → mouseup → click` (a bare `.click()` does nothing);
Outlook's fields are comboboxes that commit a draft value only on Enter *while
focused*, so `reactSet` alone silently saves the old time; and when searching
for those fields, our own row has to be excluded, because a native
`input[type=date]` reports the same `YYYY-MM-DD` shape and wins on DOM order.

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

- A userscript manager
- Python 3.11+ for the theme feed (needs `tomllib`) — not needed for the
  userscript alone
- Omarchy, for the theme feed only

## Licence

MIT
