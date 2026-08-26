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
