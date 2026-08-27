# Working on owa-minimal

A userscript that reduces Outlook Web to a calendar. It drives someone else's
React app through the DOM, so almost every rule below exists because guessing
about that DOM produced a bug that shipped.

## What this is for

Two constraints decide the design. They are not preferences to trade away when
something turns out to be hard.

**Keyboard navigation is fundamental.** The whole point is creating and moving
around events without reaching for the mouse: `j`/`k` for the week, arrow keys
for the slot, `c`, type, Tab, Enter. A change that leaves the mouse as the only
way to do something has broken the feature, whatever it does for the layout.
Judge a change by whether the keyboard path still works end to end, not by
whether the element is present.

This is not hypothetical. Reducing the box shrank the date picker's anchor,
which made its popover scroll sideways — so tabbing to End time scrolled Start
date out of sight. Everything was "there"; the keyboard path was broken. An
earlier version put the date row in the tab order when pressing Enter on it did
nothing, which is worse than leaving it out.

**Minimal, but not at the cost of the functions we want.** Removing chrome is
the means; the end is a calendar that does what a calendar has to do — create
an event, set its date and time, accept and decline invitations. When a
reduction and a capability collide, the capability wins and the reduction gets
reworked.

Also not hypothetical. Date editing was once dropped from the quick-add box
because driving Outlook's picker looked impossible — a smaller box that could
no longer set a date. That was the wrong call: the picker had been opening all
along and the strip was hiding it. If a capability seems to require giving up
the reduction, suspect the diagnosis before accepting the trade.

## Do not replace Outlook's keyboard model — scope around it

Inside the calendar surface (`[role="main"]`) Outlook already implements a
proper roving-focus model: the grid slot takes the arrow keys, and Tab steps
through the events so Enter opens one and Delete removes it. It is better than
anything worth bolting on. Reductions stop at that boundary.

Learned by breaking it. A pass that took ~60 stray controls out of the tab
order also untabbed the events, which silently removed opening and deleting
events from the keyboard; and a focus-restorer that fired whenever the active
element was not the grid slot *exactly* handed focus back the instant you
tabbed to an event, so events could not be reached at all. Both looked
reasonable and both cost a capability the reduction had no business touching.

Put focus back only when it is genuinely lost — `body`, or outside the surface
entirely — and never take it from somewhere the user deliberately put it.

## Snapshot before you paint

`takeSnapshot()` must capture **Outlook's** colours, never ours. It reads
computed styles, so if our sheets are live it captures our own output and maps
it a second time — the theme washes out and reads as no theme at all. It
disables our two style elements for the read and restores them straight after.

This has now been the same bug twice: once in the original theme code, and
again the moment cached CSS started going in at document-start, which put
painting *before* snapshotting for the first time. Any change that moves work
earlier in the load has to be checked against it.

## The one rule the others come from

**Measure it. Do not reason about what the DOM must be doing.**

Every defect in this project so far came from inferring instead of checking,
and each cost a debugging round:

- A probe filtered candidates on `height > 0`, which reports an open, correct,
  populated picker as *absent* when the caller has hidden it. Two confident,
  wrong conclusions came out of that — including "this is impossible".
- A check asked whether the fields sat inside the **modal**. They did, and it
  was still broken: the thing clipping them was the popover, and that stays
  true after the popover scrolls its own content.
- A width was capped on the modal while the shell underneath carried its own
  `width: 1122px` from a generated class, so the cap did nothing visible.

So: **check existence separately from visibility**, and **measure against the
element that actually clips**, not the nearest one you can name.

## Verify against the real thing

`verify-quickadd.js` derives the reduction from the bottom up — hide one branch
at a time from Outlook's stock compose, check the core still works, revert
anything that breaks. **Run it before changing the strip.** Its checks are
structural and never click Save, so it creates no events.

**A harness must reproduce the real call pattern, not just the real code.**
A fix once passed in a page harness and failed in the product, because the
harness called `stripCompose()` again by hand after opening the picker —
something the product never does, since it strips exactly once per compose.
Before believing a harness result, check which calls the product actually
makes, and in what order.

To test the real script in a browser without Violentmonkey, see the injection
recipe in the project notes: minify → gzip → base64 → paste → **verify a
SHA-256 prefix** → load as a `blob:` script. OWA's CSP blocks `eval` and blocks
`fetch` to `127.0.0.1`; `blob:` is allowed. Hash every paste — a corrupt chunk
once had exactly the right length.

## Selectors

OWA's class names are generated and its `aria-label`s are translated. Key on
ids, `data-app-section`, `data-unique-id`, `data-column-date`, and ARIA roles.
Office control numbers (`Ribbon-2532`) and Fabric's own ids (`DatePicker<n>`,
`ComboBox<n>-input`) are stable across locales; `aria-label="Save"` is not, and
is only a fallback.

Matching on a *value shape* is a trap: a field holding `2026-08-26` looks
exactly like any other field holding `2026-08-26`, including our own.

## Timing

Poll in **seconds, not milliseconds**. Outlook's date callout takes about a
second to mount. A 1.5s budget looked like "it never opens" often enough to
send this down the wrong path for a long time. Three seconds is the floor.

## Failure modes

**Never let a write fail silently.** If a value does not land, Outlook saves at
its own prefilled slot with no error anywhere — an event quietly filed at the
wrong time. Read the value back before allowing a save, and abandon the save if
it disagrees. A visible failure beats a silent wrong one.

## House style

- Bump `@version` on every behaviour change; the local server serves the file
  and Violentmonkey updates from it.
- Delete code that no longer works rather than leaving it behind a flag. Git
  keeps it; a file full of dead paths misleads the next reader.
- Commit messages record **what was measured** and what was wrong before, not
  just what changed. They are the project's memory — several of the findings
  above are only recoverable from them.
- `node --check` before committing. Sweep for unused identifiers after deleting.
- **Restoring code from history? List what it calls, and check each one exists.**
  `node --check` passes on a call to a function deleted twenty commits ago —
  it is a runtime `ReferenceError`, and in a userscript it surfaces as the
  feature silently doing nothing. Reinstating the quick-add fields cost three
  crashes this way (`dateRow`, `localDate`, `reactSet`), one browser round-trip
  each. The check that works is dull and manual:

  ```bash
  for id in helperOne helperTwo; do
    grep -qE "^  (const|function|let) $id\b" owa-minimal.user.js || echo "MISSING: $id"
  done
  ```

  Two attempts at a general scanner both produced ~20 false positives and gave
  *identical* output on a healthy file and a deliberately broken one, so both
  were thrown away. A check that cannot tell the two apart is worse than none —
  test any checker against a known-broken input before trusting it.
