// Paste into the browser console on an OWA calendar tab, with a compose open.
//
// Derives the quick-add reduction from the bottom up instead of assuming it:
// start from Outlook's untouched compose, hide one branch at a time, and after
// each removal check that the parts we actually need still work. Anything that
// breaks a check is put back.
//
//   Q.baseline()   what the stock compose looks like, before touching anything
//   Q.candidates() every branch that holds none of the essential nodes
//   Q.reduce()     the removal loop — returns what was kept and what was reverted
//   Q.size(w, o)   apply the shipping CSS at width w (o = picker-open width)
//   Q.verify()     one pass of the checks, including opening and closing the picker
//   Q.restore()    put everything back
//
// Why this exists: every bug in this feature so far came from reasoning about
// the DOM rather than measuring it — a `height > 0` filter that reported an
// open picker as absent, a strip that ran once when the fix needed it to run
// again, a width capped on the modal while the shell carried its own. The loop
// below cannot make those mistakes, because it only believes what it measures.
//
// Last run 2026-08-26 against Outlook's compose: 19 candidates, 0 reverted, and
// the sizes below reproduced what the userscript ships (360 closed / 560 open).

window.Q = (() => {
  const Q = {};
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const q = s => document.querySelector(s);

  Q.PICKER = 'input[id^="DatePicker"], input.ms-ComboBox-Input';
  Q.modal = () => q('[id^="ModalFocusTrapZone"]');
  Q.form = () => q('[data-app-section="Form_Content"]');
  Q.title = () => q('[id$="_SUBJECT"] input');
  Q.summary = () => {
    const dt = q('[id$="_DATETIME"]');
    return dt && [...dt.querySelectorAll('[role="button"]')]
      .find(b => /\d{4}-\d{2}-\d{2}/.test(b.textContent || ''));
  };
  Q.save = () => {
    const vis = e => e.getBoundingClientRect().height > 0;
    return [...document.querySelectorAll('button[priorityid="3"], button[overfloworderid="-3"]')].filter(vis)[0]
      || [...document.querySelectorAll('button')].filter(vis)
           .find(b => /^save$/i.test((b.getAttribute('aria-label') || '').trim()));
  };
  const vis = e => !!e && e.getBoundingClientRect().height > 0;

  // Fluent's date summary listens on pointer events; a bare .click() does
  // nothing. A second one closes the picker again, which is what makes the
  // loop below repeatable.
  const pointerClick = el => {
    const r = el.getBoundingClientRect();
    const o = { bubbles: true, cancelable: true, button: 0, pointerId: 1, isPrimary: true,
                clientX: Math.round(r.x + Math.min(40, r.width / 2)),
                clientY: Math.round(r.y + r.height / 2) };
    for (const t of ['pointerover', 'pointerenter', 'pointerdown', 'mousedown',
                     'pointerup', 'mouseup', 'click']) {
      el.dispatchEvent(t.startsWith('pointer') ? new PointerEvent(t, o) : new MouseEvent(t, o));
    }
  };

  Q.fields = () => [...document.querySelectorAll(Q.PICKER)].filter(vis);
  Q.pickerOpen = () => Q.fields().length >= 3;

  // Poll in seconds, not milliseconds: the callout takes about a second to
  // mount, and a budget shorter than that reads as "it never opened".
  Q.openPicker = async () => {
    if (Q.pickerOpen()) return true;
    const s = Q.summary();
    if (!s) return false;
    pointerClick(s);
    for (let i = 0; i < 30; i++) { if (Q.pickerOpen()) return true; await sleep(100); }
    return false;
  };
  Q.closePicker = async () => {
    if (!Q.pickerOpen()) return true;
    const s = Q.summary();
    if (s) pointerClick(s);
    for (let i = 0; i < 15; i++) { if (!Q.pickerOpen()) return true; await sleep(100); }
    return false;
  };

  // Structural only — nothing here clicks Save, so no events are created.
  Q.check = () => {
    const m = Q.modal(), t = Q.title(), s = Q.summary(), sv = Q.save();
    const mr = m && m.getBoundingClientRect();
    const sr = s && s.getBoundingClientRect();
    const over = m ? [...m.querySelectorAll('*')].map(e => e.getBoundingClientRect())
      .filter(r => r.height > 0 && r.right > mr.right + 3).length : -1;
    return {
      title: !!t && vis(t),
      dateRow: !!s && sr.width >= 150 && sr.height <= 60,
      dateRowSize: sr ? `${Math.round(sr.width)}x${Math.round(sr.height)}` : null,
      save: !!sv && vis(sv),
      saveInside: (sv && mr) ? sv.getBoundingClientRect().right <= mr.right + 3 : false,
      noOverflow: over === 0,
      modal: mr ? `${Math.round(mr.width)}x${Math.round(mr.height)}` : null,
    };
  };
  Q.ok = c => c.title && c.dateRow && c.save && c.saveInside && c.noOverflow;

  Q.verify = async () => {
    window.dispatchEvent(new Event('resize'));   // floating-ui repositions on this
    await sleep(200);
    const c = Q.check();
    if (!Q.ok(c)) return { pass: false, why: 'structure', c };
    if (!await Q.openPicker()) return { pass: false, why: 'picker', c };
    const f = Q.fields();
    const mr = Q.modal().getBoundingClientRect();
    const inside = f.every(i => {
      const b = i.getBoundingClientRect();
      return b.left >= mr.left - 3 && b.right <= mr.right + 3;
    });
    const truncated = f.some(i => i.scrollWidth > i.clientWidth + 1);
    await Q.closePicker();
    return { pass: true, c, pickerInside: inside, truncated };
  };

  // Keyboard navigation is the point of the feature, so it gets its own check.
  //
  // "Is the element present?" is not enough, and missing that cost a release:
  // shrinking the box shrank the picker's anchor, the popover started scrolling
  // its own content sideways, and tabbing to End time scrolled Start date out
  // of sight. Every field was present and inside the modal the whole time.
  //
  // So this focuses each stop in turn and then asks whether *all* the others
  // are still where you can see them — and whether focusing anything left a
  // container scrolled. A stop that displaces another stop is a broken path.
  // The invariant is about the picker's three fields, because those are the
  // ones that must be usable together. The title and Save are deliberately not
  // included: focusing either dismisses the popover, so "every stop visible at
  // once" was never true and a check built on it fails on a healthy box.
  //
  // Absence counts as failure, never as success. The first version of this
  // check passed the broken case because narrowing the popover made the fields
  // invisible, `fields()` returned none, and a loop over nothing finds nothing
  // wrong — the same existence-versus-visibility trap this file warns about.
  //
  // Verified both ways against a live compose: passes as shipped, and on a
  // popover forced back to its anchor width reports "focusing End time hid
  // Start date" and "scrolled a container sideways".
  Q.keyboard = async () => {
    const problems = [];
    if (!await Q.openPicker()) return { pass: false, problems: ['picker did not open'] };
    const f = Q.fields();
    if (f.length < 3) return { pass: false, problems: [`only ${f.length} of 3 fields visible`] };

    const label = i => i.getAttribute('aria-label');
    // Measured against the popover that clips them, not the modal: a field can
    // sit inside the modal and still be scrolled out of its own container.
    const visible = i => {
      const pop = i.closest('.fui-PopoverSurface');
      if (!pop) return false;
      const b = i.getBoundingClientRect(), c = pop.getBoundingClientRect();
      return b.width > 0 && b.left >= c.left - 3 && b.right <= c.right + 3;
    };

    for (const i of f) {
      i.focus();
      await sleep(130);
      for (const other of f) {
        if (!visible(other)) problems.push(`focusing ${label(i)} hid ${label(other)}`);
      }
      if ([...Q.modal().querySelectorAll('*')].some(el => el.scrollLeft > 1)) {
        problems.push(`focusing ${label(i)} scrolled a container sideways`);
      }
    }
    return { fields: f.map(label), pass: problems.length === 0,
             problems: [...new Set(problems)] };
  };

  Q.baseline = () => Q.check();

  // Every top-most branch holding none of the three nodes we need. Note it
  // does not descend into an essential node: the date summary's own label is
  // a child of it, and "removable" must not include that.
  Q.candidates = () => {
    const modal = Q.modal();
    const essential = [Q.title(), Q.summary(), Q.save()].filter(Boolean);
    const titleBar = [...modal.querySelectorAll('*')].find(e =>
      e.getBoundingClientRect().height > 20 &&
      [...e.querySelectorAll('button')].some(b => /close/i.test(b.getAttribute('aria-label') || '')) &&
      !e.contains(Q.title()));
    const out = [];
    const walk = el => {
      for (const c of el.children) {
        const r = c.getBoundingClientRect();
        if (r.height < 1 || r.width < 2) continue;          // already invisible
        if (c === titleBar) continue;                        // keeps close / pop-out
        if (essential.some(n => c === n || n.contains(c))) continue;
        if (essential.some(n => c.contains(n))) { walk(c); continue; }
        out.push(c);
      }
    };
    walk(modal);
    Q._candidates = out;
    Q._titleBar = titleBar;
    return out;
  };

  const style = () => {
    let st = document.getElementById('owa-verify');
    if (!st) { st = document.createElement('style'); st.id = 'owa-verify'; document.head.appendChild(st); }
    return st;
  };

  Q.reduce = async () => {
    const list = Q._candidates || Q.candidates();
    style().textContent = `[data-exp-hide]{display:none !important}`;
    const kept = [], reverted = [];
    for (const el of list) {
      const label = (el.textContent || '').trim().replace(/\s+/g, ' ').slice(0, 30)
        || `<${el.tagName}>`;
      el.setAttribute('data-exp-hide', '');
      const v = await Q.verify();
      if (v.pass) kept.push(label);
      else { el.removeAttribute('data-exp-hide'); reverted.push({ label, why: v.why }); }
    }
    return { kept, reverted, final: await Q.verify() };
  };

  // The sizing the userscript ships. Removing content does not resize the
  // modal: the shell carries its own pixel width, and the height chain has to
  // be told to collapse.
  Q.size = async (closed = 360, open = 560) => {
    const modal = Q.modal();
    const scroll = [...modal.children].find(c => c.contains(Q.form()));
    const shell = scroll && scroll.firstElementChild;
    if (shell) shell.setAttribute('data-exp-shell', '');
    let n = Q.form();
    while (n) { n.setAttribute('data-exp-fit', ''); if (n === modal) break; n = n.parentElement; }
    const w = Q.pickerOpen() ? open : closed;
    style().textContent = `[data-exp-hide]{display:none !important}
      [data-exp-shell]{width:100% !important; max-width:100% !important; min-width:0 !important}
      [data-exp-fit]{height:auto !important; min-height:0 !important;
                     flex-grow:0 !important; overflow:visible !important}
      [id^="ModalFocusTrapZone"]{width:auto !important; min-width:0 !important;
        max-width:${w}px !important; ${Q.pickerOpen() ? `min-width:${open}px !important;` : ''}
        position:relative !important; padding-bottom:56px !important}
      .ms-DatePicker, .ms-DatePicker .ms-TextField, input[id^="DatePicker"]{
        min-width:128px !important; width:auto !important}`;
    await sleep(350);
    return Q.check();
  };

  Q.restore = () => {
    const st = document.getElementById('owa-verify');
    if (st) st.remove();
    for (const a of ['data-exp-hide', 'data-exp-shell', 'data-exp-fit']) {
      for (const el of document.querySelectorAll(`[${a}]`)) el.removeAttribute(a);
    }
  };

  return Q;
})();
