// Service worker. Three jobs, all things a content script cannot do itself:
// fetch the palette feed from localhost (the page's CSP would block it),
// open the calendar as a chromeless window from the toolbar button, and
// re-inject the content script when Chrome skipped it.

const CALENDAR_URL = 'https://outlook.cloud.microsoft/calendar/view/workweek';
const CALENDAR_TABS = [
  'https://outlook.cloud.microsoft/calendar/*',
  'https://outlook.office.com/calendar/*',
  'https://outlook.office365.com/calendar/*',
];
const OUTLOOK_TABS = [
  'https://outlook.cloud.microsoft/*',
  'https://outlook.office.com/*',
  'https://outlook.office365.com/*',
];
const OUTLOOK_HOST = /^https:\/\/outlook\.(cloud\.microsoft|office\.com|office365\.com)\//;

// A declared content script is not guaranteed on the first navigation of a
// cold-started browser: launching an app window with the browser closed
// missed it in 2 of 4 tries here, and that is exactly how an installed web
// app gets opened. So every finished Outlook load is checked for the shell's
// marker and gets the scripts if it is missing. Late, and it shows the full
// Outlook for a moment first — but only on the launches that would otherwise
// have shown it for good.
async function ensureInjected(tabId) {
  try {
    const [probe] = await chrome.scripting.executeScript({
      target: { tabId },
      func: () => !!globalThis.__owaMinimalShell,
    });
    if (probe && probe.result) return;
    await chrome.scripting.executeScript({
      target: { tabId },
      files: ['core.js', 'content.js'],
    });
  } catch (_) {
    // a tab that went away, or a page we may not script — nothing to do
  }
}

chrome.tabs.onUpdated.addListener((tabId, info, tab) => {
  if (info.status !== 'complete' || !tab.url || !OUTLOOK_HOST.test(tab.url)) return;
  ensureInjected(tabId);
});

// The worker itself can start after the first page has already finished
// loading, in which case the listener above saw nothing.
chrome.tabs.query({ url: OUTLOOK_TABS })
  .then(tabs => tabs.forEach(tab => ensureInjected(tab.id)))
  .catch(() => {});

// Launch-or-focus. A second window would be a second OWA boot, which is the
// cost this whole thing exists to avoid — so an existing app window is
// brought forward instead. Only chromeless windows count; a calendar open
// in an ordinary tab is the user's tab, not the app.
chrome.action.onClicked.addListener(async () => {
  const tabs = await chrome.tabs.query({ url: CALENDAR_TABS });
  for (const tab of tabs) {
    const win = await chrome.windows.get(tab.windowId).catch(() => null);
    if (!win || (win.type !== 'popup' && win.type !== 'app')) continue;
    await chrome.windows.update(win.id, { focused: true });
    await chrome.tabs.update(tab.id, { active: true });
    return;
  }
  // Measured: a popup created this way reports display-mode: standalone,
  // the same as an installed web app, so the content script treats it as
  // the app without any flag in the URL.
  await chrome.windows.create({
    url: CALENDAR_URL,
    type: 'popup',
    width: 1280,
    height: 900,
  });
});

chrome.runtime.onMessage.addListener((msg, _sender, respond) => {
  if (!msg || msg.type !== 'fetch-theme') return false;
  fetchTheme(msg.url).then(respond, () => respond({ ok: false }));
  return true;   // answered asynchronously
});

// The feed host is an optional permission, granted from the options page
// when the user picks it. Without it the fetch would throw; checking first
// keeps the worker's console quiet for everyone who never uses a feed.
async function fetchTheme(url) {
  let origin;
  try {
    const u = new URL(url);
    origin = `${u.protocol}//${u.hostname}/*`;
  } catch (_) {
    return { ok: false, reason: 'bad-url' };
  }
  if (!(await chrome.permissions.contains({ origins: [origin] }))) {
    return { ok: false, reason: 'permission' };
  }
  const ctl = new AbortController();
  const timer = setTimeout(() => ctl.abort(), 2000);
  try {
    const res = await fetch(url, { signal: ctl.signal, cache: 'no-store' });
    if (!res.ok) return { ok: false, reason: `http-${res.status}` };
    return { ok: true, data: await res.json() };
  } catch (_) {
    return { ok: false, reason: 'unreachable' };
  } finally {
    clearTimeout(timer);
  }
}
