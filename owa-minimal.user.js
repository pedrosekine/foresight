// ==UserScript==
// @name         Outlook Web — minimal calendar (Omarchy)
// @namespace    omarchy
// @version      3.1.0
// @description  Strips OWA chrome, compresses the day scale, rebuilds a minimal action bar, and retints the whole app to the current Omarchy theme.
// @license      MIT
// @updateURL    http://127.0.0.1:8787/owa-minimal.user.js
// @downloadURL  http://127.0.0.1:8787/owa-minimal.user.js
// @match        https://outlook.office.com/*
// @match        https://outlook.office365.com/*
// @match        https://outlook.cloud.microsoft/*
// @run-at       document-start
// @grant        GM_xmlhttpRequest
// @grant        GM.xmlHttpRequest
// @grant        GM_getValue
// @grant        GM_setValue
// @connect      127.0.0.1
// @connect      localhost
// ==/UserScript==
//
// Publishing checklist — fill in once this lives in a repo, so installs
// pick up fixes automatically (OWA's DOM moves, and its hostname already
// changed once mid-development):
//   @homepageURL https://github.com/<user>/<repo>
//   @updateURL   https://raw.githubusercontent.com/<user>/<repo>/main/owa-minimal.user.js
//   @downloadURL https://raw.githubusercontent.com/<user>/<repo>/main/owa-minimal.user.js
// Left out deliberately rather than pointed at a URL that doesn't exist yet:
// a broken @updateURL makes the manager retry a 404 forever.

(function () {
  'use strict';

  // ---- scope ----------------------------------------------------------
  // A userscript matches on URL and can't tell which window it's in, so
  // without this it would reshape every Outlook tab in the browser — not
  // just the pinned app.
  //
  // The launcher appends ?omarchy=1. That flag is captured at
  // document-start (before OWA's own scripts run and rewrite the URL) into
  // sessionStorage, which is scoped to a single tab: the app window keeps
  // it across every in-app navigation, and an Outlook tab opened normally
  // never has it. Append the flag by hand to opt a regular tab in.
  const APP_FLAG = 'omarchy';
  const APP_KEY = 'omarchy-owa-app';

  function isAppWindow() {
    try {
      if (new URLSearchParams(location.search).has(APP_FLAG)) {
        sessionStorage.setItem(APP_KEY, '1');
      }
      return sessionStorage.getItem(APP_KEY) === '1';
    } catch (_) {
      return false;  // private mode with storage blocked — stay out of the way
    }
  }

  if (!isAppWindow()) return;

  // ---- config ---------------------------------------------------------
  // Editable without touching this file: Violentmonkey's script editor has
  // a "Values" tab listing these once they've been seeded below.
  const DEFAULTS = {
    hoursVisible: 14,                                 // hours on screen at once
    startHour: 6,                                     // where the grid sits on load
    themeUrl: 'http://127.0.0.1:8787/theme.json',     // palette feed from ./setup
    themePollMs: 3000,
  };

  // Wrapped rather than referenced directly — detaching these from the GM
  // object loses `this` in some managers.
  const store = {
    get: (typeof GM_getValue === 'function') ? (k => GM_getValue(k)) : null,
    set: (typeof GM_setValue === 'function') ? ((k, v) => GM_setValue(k, v)) : null,
  };
  // Seeding puts the defaults in the manager's Values tab so they can be
  // edited — but a stored value then beats the code default forever, so a
  // better default would never reach anyone who had already installed.
  // The seeded value is recorded alongside: if what is stored is still that
  // untouched seed and the default has since moved, the new one wins.
  // Anything actually edited is left alone.
  const SEED_PREFIX = '_seeded:';

  // Defaults shipped by earlier versions, from before the seed was
  // recorded. A stored value matching one of these is taken as untouched;
  // anything else is treated as deliberate and never overwritten.
  const SUPERSEDED = { startHour: [7] };

  function setting(key) {
    const fallback = DEFAULTS[key];
    if (!store.get) return fallback;
    try {
      const stored = store.get(key);
      const seed = value => {
        if (!store.set) return;
        store.set(key, value);
        store.set(SEED_PREFIX + key, value);
      };

      if (stored === undefined || stored === null || stored === '') {
        seed(fallback);
        return fallback;
      }

      const seeded = store.get(SEED_PREFIX + key);
      const noRecord = seeded === undefined || seeded === null;
      const untouched = noRecord
        ? (SUPERSEDED[key] || []).some(v => String(v) === String(stored))
        : String(seeded) === String(stored);

      if (untouched && String(stored) !== String(fallback)) {
        seed(fallback);
        return fallback;
      }
      if (noRecord && store.set) store.set(SEED_PREFIX + key, stored);

      if (typeof fallback === 'number') {
        const n = Number(stored);
        return Number.isFinite(n) ? n : fallback;
      }
      return stored;
    } catch (_) { return fallback; }
  }

  const HOURS_VISIBLE = setting('hoursVisible');
  const START_HOUR = setting('startHour');
  const THEME_URL = setting('themeUrl');
  const THEME_POLL_MS = setting('themePollMs');

  const HOTKEY = e => e.altKey && e.shiftKey && (e.key === 'O' || e.key === 'o');

  // When the Omarchy theme's mode differs from the one OWA is rendering,
  // flipping luminance makes the app follow the mode — but it also breaks
  // contrast against event category colours, which are deliberately left
  // alone, and those chips become unreadable. Mapping monotonically instead
  // preserves every contrast relationship OWA already had; the app keeps
  // OWA's light/dark but takes the palette's hues. For a real light/dark
  // switch, change OWA's own Appearance setting to match your theme.
  const INVERT_ON_MODE_MISMATCH = false;

  // Quick add uses its own date/time row instead of Outlook's, because
  // Outlook's real fields only exist inside a callout that closes on any
  // outside interaction — they cannot be tabbed through.
  const QUICK_ADD_FIELDS = true;

  const q = s => document.querySelector(s);

  // The date-navigation row. Its own data-app-section is more precise than
  // matching any [role=toolbar] inside the module — the ribbon's tab strip
  // is also a toolbar. Older builds lack the attribute, hence the fallback.
  const toolbar = () =>
    q('[data-app-section="CalendarSurfaceNavigationToolbar"]') ||
    q('[data-app-section="CalendarModule"] [role="toolbar"]');

  // Buttons lifted out of the hidden Ribbon. Each is proxied rather than
  // moved: relocating a node out of OWA's React tree gets undone on the
  // next render, but clicking the original always works.
  //
  // Located by data-unique-id, which carries the Office ribbon control
  // number and is the same in every locale. aria-label is translated, so
  // matching on it would break for anyone not running OWA in English, and
  // plain id is absent on the split buttons (New, Day). The `button`
  // qualifier matters twice over: those two also have a wrapping <div>
  // with the same attribute, and it keeps the "-Menu" chevron out.
  const ACTIONS = [
    { label: 'New', title: 'New event', primary: true, ribbon: 2532, side: 'left' },
    { label: 'D',   title: 'Day',   ribbon: 2504, side: 'right' },
    { label: 'W',   title: 'Week',  ribbon: 2519, side: 'right' },
    { label: 'M',   title: 'Month', ribbon: 2505, side: 'right' },
  ];

  const findControl = a =>
    q(`button[data-unique-id="Ribbon-${a.ribbon}"]`) ||
    q(`button[id="${a.ribbon}"]`) ||
    q(`button[aria-label="${a.title}"]`);

  // Today / previous / next, in that order. Positional because their only
  // distinguishing attribute is a translated aria-label ("Go to today").
  function navButton(i) {
    const tb = toolbar();
    if (!tb) return null;
    const btns = tb.querySelectorAll('button');
    return btns.length >= 3 ? btns[i] : null;
  }

  // ---- quick add ------------------------------------------------------
  // Outlook's compose can't be typed into: the date/time row is a
  // <div role="button">, not a field. But clicking it expands into real
  // inputs, so the whole form is drivable — no deeplink, no navigation,
  // no iframe, and nothing lost from the current view.
  const isVisible = el => !!el && el.getBoundingClientRect().height > 0;
  const composeOpen = () => !!q('[id^="EVENT_CalendarCompose"]');

  const localDate = (d = new Date()) => {
    const p = n => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
  };

  // Poll rather than observe: each step here is a React render away, and a
  // short bounded wait is easier to reason about than a tree of observers.
  function whenReady(check, cb, tries = 40) {
    const found = check();
    if (found) return cb(found);
    if (tries <= 0) return console.warn('[owa-minimal] quick add: timed out');
    setTimeout(() => whenReady(check, cb, tries - 1), 100);
  }

  // React keeps its own copy of an input's value and ignores direct
  // assignment, so go through the prototype's native setter and then fire
  // the events it listens for.
  function reactSet(input, value) {
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
    setter.call(input, value);
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
  }

  // Outlook's date and time fields are comboboxes: they hold a draft value
  // and only commit it on Enter — and only while focused. reactSet alone
  // leaves the input *showing* the new value while the form keeps the old
  // one, which then saves silently at the wrong time. Focus first, Enter
  // after; both are required.
  function commitField(input, value) {
    input.focus();
    reactSet(input, value);
    for (const type of ['keydown', 'keyup']) {
      input.dispatchEvent(new KeyboardEvent(type, {
        key: 'Enter', bubbles: true, cancelable: true, keyCode: 13, which: 13,
      }));
    }
  }

  // Which day the new event should land on. OWA always prefills today, and
  // ignores the week you're looking at — so only override when today isn't
  // on screen. Returning null means "leave OWA's own prefill alone", which
  // already gives today at the current time for 30 minutes.
  function targetDate() {
    const cols = [...document.querySelectorAll('[data-column-date]')]
      .map(e => e.dataset.columnDate).filter(Boolean).sort();
    if (!cols.length) return null;
    return cols.includes(localDate()) ? null : cols[0];
  }

  // Save carries no id, but it does carry Office toolbar metadata that is
  // the same in every locale: it is the only button in the compose with
  // priorityid="3" / overfloworderid="-3". The English label is kept as a
  // later fallback, and the last resort is positional — the first visible
  // button in the command bar, which is Save in every layout seen so far.
  function saveButton() {
    const vis = els => [...els].filter(isVisible);
    const byMeta = vis(document.querySelectorAll(
      'button[priorityid="3"], button[overfloworderid="-3"]'))[0];
    if (byMeta) return byMeta;

    const byLabel = vis(document.querySelectorAll('button')).find(b =>
      /^save$/i.test((b.getAttribute('aria-label') || '').trim())
      || /^save$/i.test((b.textContent || '').trim()));
    if (byLabel) return byLabel;

    const bar = q('[data-owa-savebar]');
    return bar ? vis(bar.querySelectorAll('button'))[0] : null;
  }

  // Fluent's date summary listens on pointer events, not click. A bare
  // .click() on it does nothing at all — this is what opens the callout
  // holding the only real date/time fields Outlook has.
  function pointerClick(el) {
    const r = el.getBoundingClientRect();
    const o = { bubbles: true, cancelable: true, button: 0, pointerId: 1, isPrimary: true,
                clientX: Math.round(r.x + Math.min(40, r.width / 2)),
                clientY: Math.round(r.y + r.height / 2) };
    for (const type of ['pointerover', 'pointerenter', 'pointerdown', 'mousedown',
                        'pointerup', 'mouseup', 'click']) {
      el.dispatchEvent(type.startsWith('pointer') ? new PointerEvent(type, o) : new MouseEvent(type, o));
    }
  }

  const dateRow = () => {
    const dt = q('[id$="_DATETIME"]');
    return dt && [...dt.querySelectorAll('[role="button"]')]
      .find(b => /\d{4}-\d{2}-\d{2}/.test((b.textContent || '').trim()));
  };

  // Outlook renders its summary as "Wed 2026-08-26 12:30 - 13:00". The date
  // is ISO and the times are 24h, so both survive without locale parsing;
  // if the shape ever changes we fall back to computing them ourselves.
  function currentSlot() {
    const row = dateRow();
    const text = row ? row.textContent : '';
    const date = (text.match(/\d{4}-\d{2}-\d{2}/) || [])[0];
    const times = text.match(/\d{1,2}:\d{2}/g) || [];
    const now = new Date();
    const pad = n => String(n).padStart(2, '0');
    const start = times[0] || `${pad(now.getHours())}:${pad(now.getMinutes())}`;
    const [h, m] = start.split(':').map(Number);
    const endDefault = `${pad((h + (m + 30 >= 60 ? 1 : 0)) % 24)}:${pad((m + 30) % 60)}`;
    return { date: date || localDate(), start, end: times[1] || endDefault };
  }

  // Our own date/time row. Outlook's real fields live in a callout that
  // closes the moment anything else is touched, so they can't be tabbed
  // through — these stand in for them and are written across on save.
  // Native date/time inputs bring a picker and keyboard entry for free.
  let qaRow = null;

  function buildFields() {
    const slot = currentSlot();
    const target = targetDate();
    qaRow = document.createElement('div');
    qaRow.id = 'omarchy-qa-row';
    const field = (type, value) => {
      const i = document.createElement('input');
      i.type = type;
      i.value = value;
      return i;
    };
    const date = field('date', target || slot.date);
    const start = field('time', slot.start);
    const end = field('time', slot.end);
    const dash = document.createElement('span');
    dash.textContent = '–';
    qaRow.append(date, start, dash, end);
    qaRow._fields = { date, start, end };

    // Keeping the end time the same length as the start when the start
    // moves, which is what you'd expect when nudging an event earlier.
    start.addEventListener('change', () => {
      const mins = s => { const [h, m] = s.split(':').map(Number); return h * 60 + m; };
      const delta = mins(end.value) - mins(slot.start);
      const total = (mins(start.value) + (delta > 0 ? delta : 30)) % 1440;
      end.value = `${String(Math.floor(total / 60)).padStart(2, '0')}:${String(total % 60).padStart(2, '0')}`;
    });
    return qaRow;
  }

  // Write our values into Outlook's callout, then save. The callout has to
  // be opened first because those inputs do not exist until it is.
  function commitAndSave() {
    const row = dateRow();
    const fields = qaRow && qaRow._fields;
    const save = saveButton();
    if (!save) return console.warn('[owa-minimal] quick add: no Save button');
    if (!row || !fields) return save.click();

    pointerClick(row);
    whenReady(
      () => {
        const inputs = [...document.querySelectorAll('input')].filter(isVisible);
        const d = inputs.find(i => /^\d{4}-\d{2}-\d{2}$/.test(String(i.value || '')));
        const t = inputs.filter(i => /^\d{1,2}:\d{2}$/.test(String(i.value || '')));
        return (d && t.length >= 2) ? { d, start: t[0], end: t[1] } : null;
      },
      () => {
        // Re-read between each: committing one field re-renders the row and
        // replaces the other inputs. Start goes before end because setting
        // the start drags the end along to keep the duration.
        const pick = () => {
          // Our own row must be excluded: a native <input type="date">
          // reports exactly YYYY-MM-DD and type="time" exactly HH:MM, so it
          // matches the same test as Outlook's fields — and being earlier in
          // the DOM, it wins. That silently commits our fields to themselves
          // and leaves Outlook's untouched, saving at the original time.
          const ins = [...document.querySelectorAll('input')]
            .filter(isVisible)
            .filter(i => !i.closest('#omarchy-qa-row'));
          return {
            d: ins.find(i => /^\d{4}-\d{2}-\d{2}$/.test(String(i.value || ''))),
            t: ins.filter(i => /^\d{1,2}:\d{2}$/.test(String(i.value || ''))),
          };
        };
        const steps = [
          () => { const p = pick(); if (p.d) commitField(p.d, fields.date.value); },
          () => { const p = pick(); if (p.t[0]) commitField(p.t[0], fields.start.value); },
          () => { const p = pick(); if (p.t[1]) commitField(p.t[1], fields.end.value); },
        ];
        steps.forEach((step, i) => setTimeout(step, i * 250));
        setTimeout(() => (saveButton() || save).click(), steps.length * 250 + 250);
      },
      30);
  }

  // Reduce the compose to the title and date rows. Walks from the title up
  // to Form_Content marking every sibling branch that leads to neither
  // keeper — which takes out the calendar picker, attendees, location, the
  // Teams toggle, the body editor and the preview pane in one pass, without
  // depending on how deeply any of them happen to be nested.
  const KEEP_ROWS = ['[id$="_SUBJECT"]', '[id$="_DATETIME"]'];
  const leadsToKeeper = el =>
    KEEP_ROWS.some(k => (el.matches && el.matches(k)) || (el.querySelector && el.querySelector(k)));

  function stripCompose() {
    const subject = q('[id$="_SUBJECT"]');
    const form = subject && subject.closest('[data-app-section="Form_Content"]');
    if (!form) return false;

    for (let node = subject; node && node !== form; node = node.parentElement) {
      for (const sib of node.parentElement.children) {
        if (!leadsToKeeper(sib)) sib.setAttribute('data-owa-hide', '');
      }
    }

    const modal = form.closest('[id^="ModalFocusTrapZone"]');
    for (let el = form; el; el = el.parentElement) {
      el.setAttribute('data-owa-fit', '');
      if (el === modal) break;
    }

    // Save alone, bottom right. The bar has to be the row that is a real
    // sibling of the form — marking an inner group instead leaves the
    // original toolbar strip behind as an empty box.
    const save = saveButton();
    if (save) {
      save.setAttribute('data-owa-savebtn', '');
      let bar = save;
      while (bar.parentElement && bar.parentElement !== form.parentElement) bar = bar.parentElement;
      bar.setAttribute('data-owa-savebar', '');
      for (const child of bar.children) {
        if (!child.contains(save)) child.setAttribute('data-owa-hide', '');
      }
    }

    // Outlook's date row is only parked away when our own fields replace
    // it. Parked rather than hidden, because commitAndSave has to open its
    // callout and that needs a layout box — and made untabbable, since an
    // off-screen row still sits in the tab order otherwise.
    if (QUICK_ADD_FIELDS) {
      const dt = q('[id$="_DATETIME"]');
      const row = dt && dt.closest('div');
      if (row) {
        row.setAttribute('data-owa-dtrow', '');
        for (const el of [row, ...row.querySelectorAll('*')]) {
          if (el.tabIndex >= 0) {
            el.setAttribute('data-owa-untab', el.getAttribute('tabindex') ?? '');
            el.tabIndex = -1;
          }
        }
      }
    }

    root.setAttribute('data-owa-quickadd', '');
    return true;
  }

  // Markers are cleared rather than left lying around, so `n` always gets
  // the full form even if it reuses nodes from a quick-add compose.
  const MARKERS = ['data-owa-hide', 'data-owa-fit', 'data-owa-savebar',
                   'data-owa-savebtn', 'data-owa-dtrow'];

  function unstripCompose() {
    root.removeAttribute('data-owa-quickadd');
    for (const attr of MARKERS) {
      for (const el of document.querySelectorAll(`[${attr}]`)) el.removeAttribute(attr);
    }
    for (const el of document.querySelectorAll('[data-owa-untab]')) {
      const prev = el.getAttribute('data-owa-untab');
      if (prev === '') el.removeAttribute('tabindex'); else el.setAttribute('tabindex', prev);
      el.removeAttribute('data-owa-untab');
    }
    qaRow?.remove();
    qaRow = null;
  }

  let quickAddActive = false;

  function quickAdd() {
    const newEvent = findControl({ ribbon: 2532, title: 'New event' });
    if (!newEvent) return console.warn('[owa-minimal] quick add: no New event button');
    unstripCompose();
    quickAddActive = true;
    newEvent.click();

    whenReady(() => q('[id$="_SUBJECT"] input'), () => {
      stripCompose();
      // Inserted after the subject row rather than appended to the modal:
      // tab order follows DOM order, and appending put it behind every
      // other focusable in the form.
      if (QUICK_ADD_FIELDS) {
        const subject = q('[id$="_SUBJECT"]');
        const dt = q('[id$="_DATETIME"]');
        const chain = el => { const a = []; for (let n = el; n; n = n.parentElement) a.push(n); return a; };
        const dtChain = new Set(chain(dt));
        const container = chain(subject).find(n => dtChain.has(n));
        const subjectRow = container && [...container.children].find(c => c.contains(subject));
        if (subjectRow) subjectRow.after(buildFields());
      }
      // Re-query rather than reuse the node we waited on: stripping makes
      // React re-render the form and replace it.
      q('[id$="_SUBJECT"] input')?.focus();
    });
  }

  // Enter in the title saves. Bound to the document rather than the input,
  // because React swaps that node out whenever the form re-renders — a
  // listener attached directly to it silently ends up on a detached
  // element and never fires again.
  function onComposeEnter(e) {
    if (e.key !== 'Enter' || e.shiftKey) return;      // Shift+Enter stays a newline
    if (!quickAddActive || !e.target || !e.target.closest) return;
    // Enter saves from anywhere in the quick-add box — the title, any of
    // our fields, or the Save button once tabbed to.
    const inQuickAdd = e.target.closest('#omarchy-qa-row')
      || e.target.closest('[id^="ModalFocusTrapZone"]');
    if (!inQuickAdd) return;
    e.preventDefault();
    e.stopPropagation();
    commitAndSave();
  }

  // Single-key shortcuts. OWA's own bindings are all modifier-based —
  // Alt+N, Alt+Shift+1..4, Ctrl+P — so plain letters are free, and those
  // keep working because anything with a modifier is ignored below.
  const SHORTCUTS = {
    c: { quickAdd: true, title: 'Quick add' },
    n: { ribbon: 2532, title: 'New event' },
    d: { ribbon: 2504, title: 'Day' },
    w: { ribbon: 2519, title: 'Week' },
    m: { ribbon: 2505, title: 'Month' },
    t: { nav: 0, title: 'Today' },
    j: { nav: 2, title: 'Next' },        // vim: j moves forward
    k: { nav: 1, title: 'Previous' },    // vim: k moves back
  };

  const isTyping = el => !!el && (
    el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' ||
    el.tagName === 'SELECT' || el.isContentEditable);

  // Two traps here. OWA keeps a zero-height dialog mounted at all times, so
  // presence alone means nothing — only a laid-out one counts. And the event
  // compose is *not* a role="dialog" at all, so without the compose check a
  // stray `d` or `m` would switch the view behind an open form.
  const overlayOpen = () => composeOpen()
    || [...document.querySelectorAll('[role="dialog"]')].some(isVisible);

  function onShortcut(e) {
    if (e.ctrlKey || e.altKey || e.metaKey) return;  // leave OWA's own bindings alone
    if (isTyping(e.target) || overlayOpen()) return;
    const binding = SHORTCUTS[e.key.toLowerCase()];
    if (!binding) return;
    if (binding.quickAdd) { e.preventDefault(); quickAdd(); return; }
    const target = ('ribbon' in binding) ? findControl(binding) : navButton(binding.nav);
    if (!target) {
      console.warn('[owa-minimal] no control for', e.key, '→', binding.title);
      return;
    }
    e.preventDefault();
    target.click();
  }

  // Elements carrying the full-day column height are found by their inline
  // height rather than by class, because OWA's class names are obfuscated
  // and churn between releases. Anything taller than this inside the
  // calendar view is a 24-hour column.
  const DAY_COLUMN_MIN_PX = 900;

  const STORE_KEY = 'omarchy-owa-minimal';
  const root = document.documentElement;

  // ---- styles ---------------------------------------------------------
  const style = document.createElement('style');
  style.id = 'omarchy-owa-style';
  style.textContent = `
    html[data-owa-minimal] #OwaTitleBar,
    html[data-owa-minimal] #LeftRail {
      display: none !important;
    }

    /* Sidebar: OWA's own left pane, collapsed by default and reduced to the
       month picker plus the per-account calendar lists. */
    html[data-owa-minimal] #leftPaneContainer { display: none !important; }
    html[data-owa-minimal][data-owa-pane] #leftPaneContainer { display: block !important; }

    /* clear our bar's row, so the toggle sits above the pane, not on it */
    html[data-owa-minimal][data-owa-pane] #leftPaneScrollContainer {
      padding-top: var(--owa-bar-height, 48px) !important;
    }

    /* "Add calendar" and "Go to my booking page". Targeted by position
       rather than by label, which is translated — and not by :has(> button
       + button), which also matches every calendar row, since those are a
       name button plus an overflow button. */
    html[data-owa-minimal][data-owa-pane]
      #leftPaneContainer [role="navigation"] > div:nth-child(2) > div:first-child {
      display: none !important;
    }

    /* The Ribbon is parked off-screen rather than display:none — Fluent
       buttons ignore programmatic clicks once they have no layout box,
       which would break every proxied action above. */
    html[data-owa-minimal] [data-app-section="Ribbon"] {
      position: fixed !important;
      top: 0 !important;
      left: -10000px !important;
      width: 100vw !important;
      pointer-events: none !important;
    }

    /* reclaim the inset around the calendar module */
    html[data-owa-minimal] [data-app-section="CalendarModule"] > div > div {
      padding: 0 !important;
    }

    /* make room for our bar at both ends of the date-navigation row */
    html[data-owa-minimal] [data-app-section="CalendarSurfaceNavigationToolbar"],
    html[data-owa-minimal] [data-app-section="CalendarModule"] [role="toolbar"] {
      padding-left: var(--owa-bar-width, 0px) !important;
      padding-right: var(--owa-bar-right, 0px) !important;
    }

    /* Today reads fine as a word; the arrows next to it do not, so this is
       aimed at that one glyph rather than at .ms-Button-icon generally. */
    html[data-owa-minimal] i[data-icon-name="CalendarTodayRegular"] {
      display: none !important;
    }

    /* the per-day "My work plan" pill in the column headers */
    html[data-owa-minimal] [role="main"] button:has(i[data-icon-name^="Building"]),
    html[data-owa-minimal] [role="main"] button[aria-label*="work plan" i] {
      display: none !important;
    }

    /* compress the 24h columns; events are positioned in percentages,
       so they follow this automatically */
    html[data-owa-minimal] [data-owa-24h] {
      height: var(--owa-day-height) !important;
    }

    /* Pinned to the toolbar's own box rather than to the viewport, so our
       buttons share Today's centre line instead of sitting 8px high on
       whatever inset OWA leaves above the row. Spans the full width so the
       view switches can sit at the far end, and is click-through except on
       the buttons themselves — OWA's own controls live underneath. */
    #omarchy-owa-bar {
      position: fixed; left: 0; width: 100vw; z-index: 60;
      top: var(--owa-bar-top, 0px);
      height: var(--owa-bar-height, 48px);
      display: none; align-items: center; justify-content: space-between;
      pointer-events: none;
    }
    html[data-owa-minimal] #omarchy-owa-bar { display: flex; }
    .omarchy-owa-group { display: flex; align-items: center; }

    /* The sidebar toggle stays pinned at the far left; everything after it
       shifts by the pane's width so New keeps sitting beside Today. */
    #omarchy-owa-left { margin-left: var(--owa-pane-width, 0px); }

    /* Every control uses Fluent's default button — the same fill, stroke,
       radius and metrics as Today. Values are adopted from the live
       toolbar at runtime so this tracks whatever theme OWA is rendering;
       the fallbacks only apply if that lookup fails. */
    .omarchy-owa-btn {
      display: inline-flex; align-items: center; justify-content: center;
      height: 32px; min-width: 40px; padding: 0 14px; margin: 8px 4px;
      font-family: var(--owa-btn-family, inherit);
      font-size: var(--owa-btn-size, 14px);
      font-weight: var(--owa-btn-weight, 400); line-height: normal; text-align: center;
      border-radius: var(--owa-btn-radius, 4px); cursor: pointer;
      color: var(--owa-btn-fg, currentColor);
      background: var(--owa-btn-bg, transparent);
      border: 1px solid var(--owa-btn-stroke, color-mix(in srgb, currentColor 30%, transparent));
    }
    .omarchy-owa-btn { pointer-events: auto; }
    #omarchy-owa-start > .omarchy-owa-btn:first-child { margin-left: 12px; }
    #omarchy-owa-right .omarchy-owa-btn:last-child { margin-right: 12px; }
    .omarchy-owa-btn.primary { padding: 0 16px; }
    .omarchy-owa-btn:hover {
      background: color-mix(in srgb, var(--owa-btn-bg, transparent) 88%, var(--owa-btn-fg, currentColor));
    }

    /* Quick add strips the compose down to title + date. The rows to drop
       are marked in JS rather than matched here: they have no ids, their
       nesting depth varies, and the one obvious CSS route (a div holding
       two adjacent buttons) also matches other parts of the form. */
    html[data-owa-quickadd] [data-owa-hide] { display: none !important; }

    /* Nothing sets an explicit height — the modal is sized by flex growth,
       so it has to be told to shrink to what is left. */
    html[data-owa-quickadd] [data-owa-fit] {
      height: auto !important;
      min-height: 0 !important;
      flex-grow: 0 !important;
    }
    html[data-owa-quickadd] [id^="ModalFocusTrapZone"] {
      width: auto !important;
      min-width: 0 !important;
      max-width: 520px !important;
      position: relative !important;
      padding-bottom: 56px !important;
    }

    /* Only Save survives the command bar, moved to the bottom right. The row
       itself is collapsed rather than hidden: a display:none ancestor would
       strip Save of its layout box, and Fluent then ignores the click. */
    html[data-owa-quickadd] [data-owa-savebar] {
      height: 0 !important; min-height: 0 !important;
      padding: 0 !important; margin: 0 !important;
      overflow: visible !important; background: none !important; border: none !important;
    }
    html[data-owa-quickadd] [data-owa-savebtn] {
      position: absolute !important; right: 16px !important; bottom: 14px !important;
      z-index: 5 !important;
    }

    /* Outlook's own date row is replaced by ours, but it has to keep a
       layout box: the real fields only exist inside a callout that opens on
       a pointer sequence, and that needs something on screen to open from. */
    html[data-owa-quickadd] [data-owa-dtrow] {
      position: absolute !important;
      left: -10000px !important;
      top: 0 !important;
      width: 560px !important;
      pointer-events: none !important;
    }

    #omarchy-qa-row {
      display: flex; align-items: center; gap: 8px;
      margin: 4px 20px 0 52px; padding-bottom: 4px;
    }
    #omarchy-qa-row input {
      font-family: var(--owa-btn-family, inherit);
      font-size: var(--owa-btn-size, 14px);
      color: var(--owa-btn-fg, currentColor);
      background: transparent;
      border: 0; border-bottom: 1px solid var(--owa-btn-stroke, currentColor);
      border-radius: 0; padding: 6px 4px; outline: none;
    }
    #omarchy-qa-row input:focus-visible {
      border-bottom-color: var(--owa-btn-fg, currentColor);
      border-bottom-width: 2px; padding-bottom: 5px;
    }
    #omarchy-qa-row input[type="date"] { flex: 1 1 auto; }
    #omarchy-qa-row input[type="time"] { flex: 0 0 auto; }
    #omarchy-qa-row span { opacity: .5; }

    #omarchy-owa-toggle {
      position: fixed; right: 6px; bottom: 6px; z-index: 2147483647;
      width: 10px; height: 10px; border: 0; border-radius: 50%;
      padding: 0; cursor: pointer; opacity: .18;
      background: currentColor; transition: opacity .15s;
    }
    #omarchy-owa-toggle:hover { opacity: .9; }
  `;
  (document.head || root).appendChild(style);

  // ---- colour helpers -------------------------------------------------
  const hex = h => {
    const m = /^#?([0-9a-f]{6})$/i.exec(h.trim());
    if (!m) return null;
    const n = parseInt(m[1], 16);
    return [n >> 16 & 255, n >> 8 & 255, n & 255];
  };
  const parseColour = v => {
    v = (v || '').trim();
    if (/^#[0-9a-f]{3}$/i.test(v)) {
      const c = v.slice(1);
      return { rgb: hex('#' + c[0] + c[0] + c[1] + c[1] + c[2] + c[2]), a: 1 };
    }
    if (v[0] === '#') { const r = hex(v); return r && { rgb: r, a: 1 }; }
    const m = /^rgba?\(([^)]+)\)$/.exec(v);
    if (!m) return null;
    const p = m[1].split(/[,\s/]+/).filter(Boolean).map(Number);
    if (p.length < 3 || p.slice(0, 3).some(isNaN)) return null;
    return { rgb: p.slice(0, 3), a: p.length > 3 ? p[3] : 1 };
  };
  const lum = ([r, g, b]) => (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255;
  const isGrey = ([r, g, b]) => Math.max(r, g, b) - Math.min(r, g, b) <= 10;
  const fmt = ([r, g, b], a) => a >= 1
    ? `rgb(${Math.round(r)}, ${Math.round(g)}, ${Math.round(b)})`
    : `rgba(${Math.round(r)}, ${Math.round(g)}, ${Math.round(b)}, ${a})`;
  const toHsl = ([r, g, b]) => {
    r /= 255; g /= 255; b /= 255;
    const mx = Math.max(r, g, b), mn = Math.min(r, g, b), l = (mx + mn) / 2;
    let h = 0, s = 0;
    if (mx !== mn) {
      const d = mx - mn;
      s = l > .5 ? d / (2 - mx - mn) : d / (mx + mn);
      h = mx === r ? ((g - b) / d + (g < b ? 6 : 0)) : mx === g ? ((b - r) / d + 2) : ((r - g) / d + 4);
      h /= 6;
    }
    return [h, s, l];
  };
  const toRgb = ([h, s, l]) => {
    if (!s) return [l * 255, l * 255, l * 255];
    const q2 = l < .5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q2;
    const f = t => {
      t = (t + 1) % 1;
      return t < 1 / 6 ? p + (q2 - p) * 6 * t : t < 1 / 2 ? q2
           : t < 2 / 3 ? p + (q2 - p) * (2 / 3 - t) * 6 : p;
    };
    return [f(h + 1 / 3) * 255, f(h) * 255, f(h - 1 / 3) * 255];
  };

  // ---- theme engine ---------------------------------------------------
  // OWA runs two token systems side by side — Fluent v9 (--colorNeutral*)
  // and the older Fabric slots (--neutralPrimary, --white, --themePrimary).
  // Rather than enumerate either, every custom property whose value parses
  // as a colour is considered. Greys are projected onto the Omarchy neutral
  // ramp by luminance and brand blues take the accent's hue at their own
  // lightness, which leaves semantic colours — status reds/greens and your
  // event categories — untouched.
  const RAMP_KEYS = ['darker_background', 'dark_background', 'background', 'selection',
    'lighter_background', 'muted', 'dark_foreground', 'foreground',
    'light_foreground', 'bright_foreground'];

  // The handful of tokens that decide how the app reads at a glance; a pure
  // luminance projection lands near these but not exactly on them.
  const ANCHORS = {
    NeutralBackground1: 'background', NeutralBackground1Hover: 'lighter_background',
    NeutralBackground1Selected: 'selection', NeutralBackground2: 'dark_background',
    NeutralBackground3: 'darker_background', NeutralBackground4: 'darker_background',
    NeutralBackground5: 'darker_background', NeutralBackground6: 'lighter_background',
    NeutralForeground1: 'bright_foreground', NeutralForeground2: 'foreground',
    NeutralForeground3: 'foreground', NeutralForeground4: 'dark_foreground',
    NeutralStroke1: 'muted', NeutralStroke2: 'selection', NeutralStroke3: 'selection',
    BrandBackground: 'accent', BrandForeground1: 'accent',
    CompoundBrandBackground: 'accent', CompoundBrandStroke: 'accent',
    NeutralForegroundOnBrand: 'background',
  };

  const COLOUR_PROPS = ['color', 'background-color', 'border-color', 'border-top-color',
    'border-right-color', 'border-bottom-color', 'border-left-color', 'fill', 'stroke',
    'outline-color', 'caret-color', 'text-decoration-color', 'column-rule-color'];

  let snapshot = null;   // pristine token values, captured once
  let lastMtime = null;
  let lastSheetCount = 0;

  // Must run before any override lands, and always map from this — mapping
  // from live values would compound the remap on every theme change.
  function takeSnapshot() {
    const cs = getComputedStyle(root);
    if (!cs.getPropertyValue('--colorNeutralBackground1').trim()) return false;
    snapshot = {};
    for (let i = 0; i < cs.length; i++) {
      const p = cs[i];
      if (p.startsWith('--')) snapshot[p] = cs.getPropertyValue(p).trim();
    }
    return true;
  }

  function makeMapper(palette) {
    const C = palette.colors;
    const ramp = RAMP_KEYS.map(k => C[k]).filter(Boolean).map(hex).filter(Boolean)
      .sort((a, b) => lum(a) - lum(b));
    if (ramp.length < 2) return null;
    const lo = lum(ramp[0]), hi = lum(ramp[ramp.length - 1]);
    const project = L => {
      const t = lo + (hi - lo) * Math.min(Math.max(L, 0), 1);
      for (let i = 1; i < ramp.length; i++) {
        const a = ramp[i - 1], b = ramp[i], la = lum(a), lb = lum(b);
        if (t <= lb || i === ramp.length - 1) {
          const f = lb === la ? 0 : (t - la) / (lb - la);
          return [0, 1, 2].map(k => a[k] + (b[k] - a[k]) * Math.min(Math.max(f, 0), 1));
        }
      }
    };
    const accent = hex(C.accent) || ramp[ramp.length - 1];
    const aH = toHsl(accent);
    const brandish = rgb => { const [h, s] = toHsl(rgb); return s > 0.35 && h > 0.5 && h < 0.68; };

    const owaDark = lum(parseColour(snapshot['--colorNeutralBackground1']).rgb) < 0.5;
    const invert = INVERT_ON_MODE_MISMATCH && (palette.mode === 'dark') !== owaDark;
    const adj = L => invert ? 1 - L : L;

    // Fluent's primary accent appears verbatim in tokens, in the Fabric
    // slots (--themePrimary) and in baked-in rules alike — day numbers, the
    // now-line, the selection block. Those should land on the palette accent
    // exactly. Sending them through the hue/saturation path instead lands a
    // shade off, because that path deliberately preserves the source's own
    // lightness, and Fluent's accent is darker than most palette accents.
    // Read from the snapshot rather than hardcoded, so it stays correct if
    // Microsoft changes the shade.
    const brandRef = parseColour(snapshot['--colorBrandForeground1'] || '');
    const same = (a, b) => a[0] === b[0] && a[1] === b[1] && a[2] === b[2];

    return v => {
      const c = parseColour(v);
      if (!c) return null;
      if (isGrey(c.rgb)) return fmt(project(adj(lum(c.rgb))), c.a);
      if (brandRef && same(c.rgb, brandRef.rgb)) return fmt(accent, c.a);
      if (brandish(c.rgb)) return fmt(toRgb([aH[0], aH[1], adj(toHsl(c.rgb)[2])]), c.a);
      return null;
    };
  }

  function sheet(id) {
    let st = document.getElementById(id);
    if (!st) {
      st = document.createElement('style');
      st.id = id;
      (document.head || root).appendChild(st);
    }
    return st;
  }

  function applyTokens(palette, map) {
    const decls = [];
    for (const [prop, val] of Object.entries(snapshot)) {
      const m = map(val);
      if (m) decls.push(`${prop}:${m} !important`);
    }
    for (const [tok, key] of Object.entries(ANCHORS)) {
      if (palette.colors[key]) decls.push(`--color${tok}:${palette.colors[key]} !important`);
    }
    sheet('omarchy-owa-theme').textContent =
      `html, body, [class*="fui-FluentProvider"], [class*="ms-Fabric"] {\n${decls.join(';\n')}\n}`;
  }

  // A minority of OWA's rules bake colours in as literals instead of
  // referencing a token, so those are re-emitted as !important overrides
  // against the same selectors.
  //
  // A full pass is ~30ms here (60 sheets, ~24k rules), which is a dropped
  // frame — and Griffel keeps injecting stylesheets as components mount, so
  // repeating it on every injection would stutter. Sheets already processed
  // are therefore remembered and only new ones are walked; a full rebuild
  // happens solely when the palette itself changes.
  let seenSheets = new WeakSet();
  let literalCss = [];

  function applyLiterals(map, full) {
    if (full) { seenSheets = new WeakSet(); literalCss = []; }
    const rules = [];
    const walk = (list, wrap) => {
      for (const r of list) {
        if (r.type === CSSRule.MEDIA_RULE) { walk(r.cssRules, r.conditionText); continue; }
        if (r.type === CSSRule.SUPPORTS_RULE) { walk(r.cssRules, wrap); continue; }
        if (r.type !== CSSRule.STYLE_RULE) continue;
        const owner = r.parentStyleSheet && r.parentStyleSheet.ownerNode;
        if (owner && /^omarchy-owa-/.test(owner.id || '')) continue;
        const body = [];
        for (const p of COLOUR_PROPS) {
          const v = r.style.getPropertyValue(p);
          if (!v || v.includes('var(')) continue;
          const m = map(v);
          if (m) body.push(`${p}:${m} !important`);
        }
        if (body.length) {
          const rule = `${r.selectorText}{${body.join(';')}}`;
          rules.push(wrap ? `@media ${wrap}{${rule}}` : rule);
        }
      }
    };
    for (const ss of document.styleSheets) {
      if (seenSheets.has(ss)) continue;
      seenSheets.add(ss);
      try { walk(ss.cssRules, null); } catch (_) { /* cross-origin */ }
    }
    if (rules.length) {
      literalCss.push(...rules);
      sheet('omarchy-owa-literals').textContent = literalCss.join('\n');
    }
    lastSheetCount = document.styleSheets.length;
  }

  let palette = null;
  // full=true remaps every sheet from scratch (the palette changed);
  // full=false only picks up sheets injected since the last pass.
  function paint(full) {
    if (!palette || !snapshot) return;
    const map = makeMapper(palette);
    if (!map) return;
    if (full) applyTokens(palette, map);
    applyLiterals(map, full);
  }

  const xhr = (typeof GM_xmlhttpRequest === 'function') ? GM_xmlhttpRequest
            : (typeof GM !== 'undefined' && GM.xmlHttpRequest) ? GM.xmlHttpRequest : null;
  let warned = false;

  function pollTheme() {
    if (!xhr) {
      if (!warned) { warned = true; console.warn('[owa-minimal] no GM_xmlhttpRequest; theme sync disabled'); }
      return;
    }
    xhr({
      method: 'GET', url: THEME_URL, timeout: 2000,
      onload: res => {
        if (res.status !== 200) return;
        let data;
        try { data = JSON.parse(res.responseText); } catch (_) { return; }
        // Snapshot first, and don't record the mtime until it succeeds. The
        // first poll can land before OWA has defined its tokens; banking the
        // mtime there would make every later poll see "unchanged" and bail,
        // wedging the theme off until the next time the palette changed.
        if (!snapshot && !takeSnapshot()) return;
        if (data.mtime === lastMtime) return;   // genuinely unchanged
        lastMtime = data.mtime;
        palette = data;
        paint(true);
      },
      onerror: () => {}, ontimeout: () => {},
    });
  }

  // ---- core -----------------------------------------------------------
  const inCalendar = () =>
    !!document.querySelector('[data-app-section="CalendarModule"]');

  function tagDayColumns() {
    const main = q('[role="main"]');
    if (!main) return;
    for (const el of main.querySelectorAll('[style*="height:"]')) {
      if (el.hasAttribute('data-owa-24h')) continue;
      const m = /height:\s*(\d+(?:\.\d+)?)px/.exec(el.getAttribute('style') || '');
      if (m && parseFloat(m[1]) >= DAY_COLUMN_MIN_PX) {
        el.setAttribute('data-owa-24h', '');
      }
    }
  }

  function rescale() {
    const sc = q('.inDayScrollContainer');
    if (!sc || sc.clientHeight < 50) return;
    const h = Math.round((24 * sc.clientHeight) / HOURS_VISIBLE);
    root.style.setProperty('--owa-day-height', h + 'px');
    return sc;
  }

  // Scroll to START_HOUR once per freshly rendered grid, so it doesn't
  // fight the user after they scroll themselves.
  let scrollDeadline = 0, scrollHeld = 0;
  function settleScroll(sc) {
    if (!sc || sc.hasAttribute('data-owa-scrolled')) return;
    if (sc.scrollHeight <= sc.clientHeight) return;
    const target = Math.round((sc.scrollHeight / 24) * START_HOUR);
    if (!scrollDeadline) scrollDeadline = performance.now() + 3000;

    // OWA sets its own scroll position once the grid mounts, sometimes after
    // ours, which left the day opening at 00:00. Re-apply until it has held
    // for two passes, then stop — so the user's own scrolling is never
    // fought beyond the initial settling window.
    if (Math.abs(sc.scrollTop - target) <= 2) {
      if (++scrollHeld >= 2) sc.setAttribute('data-owa-scrolled', '');
      return;
    }
    scrollHeld = 0;
    sc.scrollTop = target;
    if (performance.now() > scrollDeadline) sc.setAttribute('data-owa-scrolled', '');
  }

  // Copy Fluent's own button tokens off the live Today button, so the
  // injected bar restyles itself whenever OWA's theme changes.
  function adoptToolbarTheme() {
    const tb = toolbar();
    if (!tb) return;
    const [def] = tb.querySelectorAll('button');
    if (!def) return;
    const a = getComputedStyle(def), s = root.style;
    s.setProperty('--owa-btn-family', a.fontFamily);
    s.setProperty('--owa-btn-size', a.fontSize);
    s.setProperty('--owa-btn-radius', a.borderTopLeftRadius);
    s.setProperty('--owa-btn-fg', a.color);
    s.setProperty('--owa-btn-bg', a.backgroundColor);
    s.setProperty('--owa-btn-stroke', a.borderTopColor);

    // Weight lives on the label, not the button box — the box reports 400
    // while Today actually renders at 600. Pick the leaf carrying the most
    // text, since the first leaf is the icon glyph.
    const label = [...def.querySelectorAll('*')]
      .filter(e => !e.children.length && e.textContent.trim())
      .sort((x, y) => y.textContent.trim().length - x.textContent.trim().length)[0];
    if (label) s.setProperty('--owa-btn-weight', getComputedStyle(label).fontWeight);
  }

  // Matches the 4px side margins our own buttons carry, so every gap in the
  // row reads the same.
  const BUTTON_GAP = 8;

  let barPad = null, barTop = null, barRight = null, paneWidth = null;
  // The bar is out of flow, so measuring the toolbar it sits over can't
  // feed back into its own position.
  function syncBar() {
    if (!bar) return;
    const tb = toolbar();
    if (!tb) return;

    const r = tb.getBoundingClientRect();
    if (r.height && r.top !== barTop) {
      barTop = r.top;
      root.style.setProperty('--owa-bar-top', r.top + 'px');
      root.style.setProperty('--owa-bar-height', r.height + 'px');
    }

    // Keep the right-hand group clear of the toolbar's own content.
    const rw = rightGroup ? rightGroup.offsetWidth : 0;
    if (rw !== barRight) {
      barRight = rw;
      root.style.setProperty('--owa-bar-right', rw + 'px');
    }

    const pane = q('#leftPaneContainer');
    const pw = pane ? pane.offsetWidth : 0;
    if (pw !== paneWidth) {
      paneWidth = pw;
      root.style.setProperty('--owa-pane-width', pw + 'px');
      barPad = null;   // the row just reflowed; re-derive the gap from scratch
    }

    // The distance from our last left-hand button to Today is the sum of our
    // margin, the toolbar's own inset and Today's 12px margin — none of which
    // are ours to predict. Measure the real gap and correct the toolbar's
    // padding by the difference; being linear, it lands in one pass.
    const today = tb.querySelector('button');
    const last = leftGroup && leftGroup.lastElementChild;
    if (!today || !last) return;
    if (barPad === null) {
      barPad = leftGroup.offsetWidth;
      root.style.setProperty('--owa-bar-width', barPad + 'px');
    }
    const gap = today.getBoundingClientRect().left - last.getBoundingClientRect().right;
    if (Math.abs(gap - BUTTON_GAP) > 0.5) {
      barPad = Math.min(Math.max(barPad + BUTTON_GAP - gap, 0), 2000);
      root.style.setProperty('--owa-bar-width', barPad + 'px');
    }
  }

  let queued = false;
  function update() {
    if (queued) return;
    queued = true;
    requestAnimationFrame(() => {
      queued = false;
      if (palette && document.styleSheets.length !== lastSheetCount) paint(false);
      if (quickAddActive && !composeOpen()) { quickAddActive = false; unstripCompose(); }
      if (!root.hasAttribute('data-owa-minimal') || !inCalendar()) return;
      adoptToolbarTheme();
      syncBar();
      tagDayColumns();
      settleScroll(rescale());
    });
  }

  const PANE_KEY = 'omarchy-owa-pane';
  function setPane(on) {
    if (on) root.setAttribute('data-owa-pane', '');
    else root.removeAttribute('data-owa-pane');
    try { localStorage.setItem(PANE_KEY, on ? '1' : '0'); } catch (_) {}
    barPad = null;
    update();
  }

  function setEnabled(on) {
    if (on) root.setAttribute('data-owa-minimal', '');
    else root.removeAttribute('data-owa-minimal');
    try { localStorage.setItem(STORE_KEY, on ? '1' : '0'); } catch (_) {}
    if (on) update();
  }

  let bar = null, leftGroup = null, rightGroup = null;
  function buildBar() {
    bar = document.createElement('div');
    bar.id = 'omarchy-owa-bar';

    const start = document.createElement('div');
    start.id = 'omarchy-owa-start';
    start.className = 'omarchy-owa-group';
    leftGroup = document.createElement('div');
    leftGroup.id = 'omarchy-owa-left';
    leftGroup.className = 'omarchy-owa-group';
    rightGroup = document.createElement('div');
    rightGroup.id = 'omarchy-owa-right';
    rightGroup.className = 'omarchy-owa-group';

    const paneToggle = document.createElement('button');
    paneToggle.className = 'omarchy-owa-btn';
    paneToggle.textContent = '☰';
    paneToggle.title = 'Calendars';
    paneToggle.addEventListener('click', () =>
      setPane(!root.hasAttribute('data-owa-pane')));

    start.append(paneToggle, leftGroup);
    bar.append(start, rightGroup);

    for (const a of ACTIONS) {
      const b = document.createElement('button');
      b.className = 'omarchy-owa-btn' + (a.primary ? ' primary' : '');
      b.textContent = a.label;
      b.title = a.title;
      b.addEventListener('click', () => {
        const target = findControl(a);
        if (target) target.click();
        else console.warn('[owa-minimal] control not found:', a.title);
      });
      (a.side === 'right' ? rightGroup : leftGroup).appendChild(b);
    }
    document.body.appendChild(bar);
  }

  // ---- wire up --------------------------------------------------------
  let enabled = true;
  try { enabled = localStorage.getItem(STORE_KEY) !== '0'; } catch (_) {}
  setEnabled(enabled);

  addEventListener('keydown', e => {
    if (!HOTKEY(e)) return;
    e.preventDefault();
    setEnabled(!root.hasAttribute('data-owa-minimal'));
  });

  // Capture phase, so a shortcut is never swallowed by a handler further in.
  addEventListener('keydown', onShortcut, true);
  addEventListener('keydown', onComposeEnter, true);

  addEventListener('resize', update);

  const start = () => {
    new MutationObserver(update).observe(document.body, {
      childList: true,
      subtree: true,
    });

    buildBar();

    let paneOpen = false;
    try { paneOpen = localStorage.getItem(PANE_KEY) === '1'; } catch (_) {}
    setPane(paneOpen);

    const btn = document.createElement('button');
    btn.id = 'omarchy-owa-toggle';
    btn.title = 'Toggle full Outlook (Alt+Shift+O)';
    btn.addEventListener('click', () =>
      setEnabled(!root.hasAttribute('data-owa-minimal')));
    document.body.appendChild(btn);

    pollTheme();
    setInterval(pollTheme, THEME_POLL_MS);
    update();
  };

  if (document.body) start();
  else addEventListener('DOMContentLoaded', start);
})();
