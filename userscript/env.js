// Userscript shell around core.js: Violentmonkey's synchronous GM_* storage
// and its cross-origin XHR, wrapped into the `env` the core expects.
//
// A userscript matches on URL and cannot tell which window it is in, so
// without the ?omarchy=1 flag it would reshape every Outlook tab in the
// browser — not just the pinned app. The launcher appends the flag; the
// core stashes it in sessionStorage so it survives in-app navigation.
(function () {
  'use strict';

  if (!foresight.flaggedWindow()) return;

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

  function setting(key, fallback) {
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

  // The userscript is the Omarchy install, so the feed is always its source.
  const THEME_URL = setting('themeUrl', 'http://127.0.0.1:8787/theme.json');

  const xhr = (typeof GM_xmlhttpRequest === 'function') ? GM_xmlhttpRequest
            : (typeof GM !== 'undefined' && GM.xmlHttpRequest) ? GM.xmlHttpRequest : null;
  let warned = false;

  function theme() {
    if (!xhr) {
      if (!warned) { warned = true; console.warn('[foresight] no GM_xmlhttpRequest; theme sync disabled'); }
      return Promise.resolve(null);
    }
    return new Promise(resolve => {
      xhr({
        method: 'GET', url: THEME_URL, timeout: 2000,
        onload: res => {
          if (res.status !== 200) return resolve(null);
          try { resolve(JSON.parse(res.responseText)); } catch (_) { resolve(null); }
        },
        onerror: () => resolve(null),
        ontimeout: () => resolve(null),
      });
    });
  }

  foresight({
    setting,
    cache: {
      get: key => (store.get ? store.get(key) : undefined),
      set: (key, value) => { if (store.set) store.set(key, value); },
    },
    theme,
  });
})();
