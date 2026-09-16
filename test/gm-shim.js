// Violentmonkey stand-ins for the userscript smoke test. The feed answers
// with a fixed palette after a tick.
window.__cache = {};
function GM_getValue(k) { return window.__cache[k]; }
function GM_setValue(k, v) { window.__cache[k] = v; }
function GM_xmlhttpRequest(o) {
  const palette = { mode: 'dark', mtime: 1, colors: {
    background: '#1a1b26', dark_background: '#13141c', darker_background: '#0e0e14',
    lighter_background: '#24283b', selection: '#292e42', muted: '#414868',
    foreground: '#a9b1d6', dark_foreground: '#565f89', light_foreground: '#b4bee6',
    bright_foreground: '#c0caf5', accent: '#7aa2f7' } };
  setTimeout(() => o.onload({ status: 200, responseText: JSON.stringify(palette) }), 10);
}
