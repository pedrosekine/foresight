// Options page. Everything lives in chrome.storage.sync under the same keys
// the content script reads; only values the user actually set are stored,
// so a better default still reaches an existing install.
(async () => {
  'use strict';

  const DEFAULTS = {
    hoursVisible: 14,
    startHour: 6,
    themeSource: 'none',
    themeUrl: 'http://127.0.0.1:8787/theme.json',
    applyInTabs: false,
    landOnCalendar: true,
  };
  const FEED_ORIGINS = ['http://127.0.0.1/*', 'http://localhost/*'];

  const $ = id => document.getElementById(id);

  // Bundled palettes, listed from their index so adding one is a file drop.
  const index = await fetch(chrome.runtime.getURL('palettes/index.json'))
    .then(r => r.json()).catch(() => []);
  for (const p of index) {
    const opt = document.createElement('option');
    opt.value = p.id;
    opt.textContent = `${p.label}${p.mode === 'light' ? ' (light)' : ''}`;
    $('bundled').appendChild(opt);
  }

  const stored = await chrome.storage.sync.get(null);
  const current = key => (stored[key] === undefined ? DEFAULTS[key] : stored[key]);

  function render() {
    $('hoursVisible').value = stored.hoursVisible ?? '';
    $('startHour').value = stored.startHour ?? '';
    $('themeSource').value = current('themeSource');
    $('themeUrl').value = stored.themeUrl ?? '';
    $('applyInTabs').checked = !!current('applyInTabs');
    $('landOnCalendar').checked = current('landOnCalendar') !== false;
    $('feedRow').hidden = current('themeSource') !== 'feed';
  }
  render();

  async function save(key, value) {
    const isDefault = value === '' || value === null || value === undefined
      || String(value) === String(DEFAULTS[key]);
    if (isDefault) {
      delete stored[key];
      await chrome.storage.sync.remove(key);
    } else {
      stored[key] = value;
      await chrome.storage.sync.set({ [key]: value });
    }
  }

  const number = (id, min, max) => async () => {
    const raw = $(id).value.trim();
    if (raw === '') return save(id, undefined);
    const n = Number(raw);
    if (!Number.isFinite(n) || n < min || n > max) { render(); return; }
    await save(id, Math.round(n));
  };
  $('hoursVisible').addEventListener('change', number('hoursVisible', 4, 24));
  $('startHour').addEventListener('change', number('startHour', 0, 23));

  $('applyInTabs').addEventListener('change', e => save('applyInTabs', e.target.checked));
  $('landOnCalendar').addEventListener('change', e => save('landOnCalendar', e.target.checked));

  // The feed lives on localhost, which is an optional host permission: ask
  // for it the moment the feed is chosen, while we still have the click.
  $('themeSource').addEventListener('change', async e => {
    const value = e.target.value;
    if (value === 'feed') {
      const ok = await chrome.permissions.request({ origins: FEED_ORIGINS });
      if (!ok) { render(); return; }
    }
    await save('themeSource', value);
    render();
    if (value === 'feed') checkFeed();
  });

  $('themeUrl').addEventListener('change', async e => {
    await save('themeUrl', e.target.value.trim());
    checkFeed();
  });

  async function checkFeed() {
    const url = current('themeUrl');
    const res = await chrome.runtime.sendMessage({ type: 'fetch-theme', url }).catch(() => null);
    $('feedStatus').textContent = res && res.ok
      ? `Feed answering: ${res.data.mode || ''} palette, ${Object.keys(res.data.colors || {}).length} colours.`
      : res && res.reason === 'permission'
        ? 'Chrome has not granted access to that host. Pick the option again to allow it.'
        : 'Feed not answering. On Omarchy, run ./setup from the repository to start it.';
  }
  if (current('themeSource') === 'feed') checkFeed();

  $('reset').addEventListener('click', async () => {
    await chrome.storage.sync.clear();
    for (const k of Object.keys(stored)) delete stored[k];
    render();
  });
})();
