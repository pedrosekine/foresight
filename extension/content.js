// Extension shell around core.js. Runs at document_start on every Outlook
// page and decides, before OWA has painted anything, whether this window is
// the calendar app or an ordinary tab that should be left alone.
//
// Nothing here touches the DOM beyond what core.js does; the job of this
// file is to turn Chrome's asynchronous storage into the synchronous `env`
// the core expects, and to pick a theme source.
(async () => {
  'use strict';

  // Chrome does not always run a declared content script on the very first
  // navigation of a cold-started browser — measured at 2 misses in 4 launches
  // of an app window here. background.js re-injects when that happens, so
  // this has to be safe to run twice: the marker lives in the extension's
  // isolated world, invisible to the page.
  if (globalThis.__foresightShell) return;
  globalThis.__foresightShell = true;

  // Measured on Chromium 152: a window opened with --app= or from an
  // installed web app reports `standalone`; a normal tab reports `browser`.
  // That is the whole "regular site untouched, app window reduced" split,
  // with no launcher flag needed. The old ?omarchy=1 flag still opts a tab
  // in, and so does the setting.
  const standalone = matchMedia('(display-mode: standalone)').matches;
  const flagged = foresight.flaggedWindow();

  const [sync, local] = await Promise.all([
    chrome.storage.sync.get(null),
    chrome.storage.local.get(null),
  ]);
  const settings = sync || {};
  const cache = local || {};

  if (!standalone && !flagged && !settings.applyInTabs) return;

  // The app is a calendar. An installed Outlook web app starts on mail, so
  // in app mode the mail landing is sent straight on to the calendar —
  // otherwise the reduction hides the rail that would get you there.
  // Only the two Outlook landings are redirected; auth hops and deep links
  // are left alone.
  if (standalone && settings.landOnCalendar !== false
      && /^\/(mail(\/.*)?)?\/?$/.test(location.pathname)) {
    location.replace('/calendar/view/workweek');
    return;
  }

  const DEFAULT_FEED = 'http://127.0.0.1:8787/theme.json';

  const env = {
    setting(key, fallback) {
      const v = settings[key];
      if (v === undefined || v === null || v === '') return fallback;
      if (typeof fallback === 'number') {
        const n = Number(v);
        return Number.isFinite(n) ? n : fallback;
      }
      return v;
    },
    cache: {
      get: key => cache[key],
      set(key, value) {
        cache[key] = value;
        chrome.storage.local.set({ [key]: value }).catch(() => {});
      },
    },
    // Resolved on every poll, so a source changed on the options page takes
    // effect at the next tick without a reload.
    async theme() {
      const source = settings.themeSource || 'none';
      if (source === 'none') return null;
      if (source === 'feed') {
        const url = settings.themeUrl || DEFAULT_FEED;
        const res = await chrome.runtime.sendMessage({ type: 'fetch-theme', url });
        return res && res.ok ? res.data : null;
      }
      // A bundled palette. Its identity is the version, so `mtime` only
      // changes when the user picks another one.
      const data = await bundledPalette(source);
      return data ? { ...data, mtime: `bundled:${source}` } : null;
    },
  };

  const palettes = new Map();
  async function bundledPalette(name) {
    if (!/^[a-z0-9-]+$/.test(name)) return null;
    if (!palettes.has(name)) {
      palettes.set(name, fetch(chrome.runtime.getURL(`palettes/${name}.json`))
        .then(r => (r.ok ? r.json() : null))
        .catch(() => null));
    }
    return palettes.get(name);
  }

  chrome.storage.onChanged.addListener((changes, area) => {
    if (area !== 'sync') return;
    for (const [key, { newValue }] of Object.entries(changes)) {
      if (newValue === undefined) delete settings[key];
      else settings[key] = newValue;
    }
  });

  foresight(env);
})();
