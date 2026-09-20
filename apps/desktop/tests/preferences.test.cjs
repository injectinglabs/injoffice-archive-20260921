const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

async function loadPreferences() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/preferences.ts'),
    platform: 'node',
    transform: { jsx: { runtime: 'automatic' } },
  });
  try {
    const { output } = await bundle.generate({ format: 'cjs', codeSplitting: false });
    const mod = { exports: {} };
    new Function('require', 'module', 'exports', output[0].code)(require, mod, mod.exports);
    return mod.exports;
  } finally {
    await bundle.close();
  }
}

function storage(seed = {}) {
  const map = new Map(Object.entries(seed));
  return {
    getItem: key => (map.has(key) ? map.get(key) : null),
    setItem: (key, value) => { map.set(key, String(value)); },
    map,
  };
}

test('preferences fail closed on damaged storage and write a versioned blob', async () => {
  const { defaultPreferences, readPreferences, writePreferences, initialView } = await loadPreferences();
  global.window = { localStorage: storage({ 'injoffice.preferences.v1': '{not json' }) };
  assert.deepEqual(readPreferences(), defaultPreferences);
  global.window = { localStorage: storage({ 'injoffice.preferences.v1': JSON.stringify({ version: 1, defaultZoom: 10, showNavigation: true }) }) };
  assert.deepEqual(readPreferences(), defaultPreferences);
  const store = storage();
  global.window = { localStorage: store };
  writePreferences({ version: 1, defaultZoom: 125, showNavigation: false, theme: 'dark' });
  assert.deepEqual(readPreferences(), { version: 1, defaultZoom: 125, showNavigation: false, theme: 'dark' });
  assert.deepEqual(initialView(readPreferences()), { zoom: 125, navigation: false, focus: false });
  assert.equal(JSON.parse(store.map.get('injoffice.preferences.v1')).version, 1);
});

test('theme preference follows the system when absent and fails closed on unknown values', async () => {
  const { defaultPreferences, readPreferences, themePreferences } = await loadPreferences();
  assert.deepEqual(themePreferences, ['system', 'light', 'dark']);
  assert.equal(defaultPreferences.theme, 'system');
  // Blobs written before the theme existed keep their view defaults and follow the system.
  global.window = { localStorage: storage({ 'injoffice.preferences.v1': JSON.stringify({ version: 1, defaultZoom: 150, showNavigation: false }) }) };
  assert.deepEqual(readPreferences(), { version: 1, defaultZoom: 150, showNavigation: false, theme: 'system' });
  for (const theme of ['light', 'dark', 'system']) {
    global.window = { localStorage: storage({ 'injoffice.preferences.v1': JSON.stringify({ version: 1, defaultZoom: 150, showNavigation: false, theme }) }) };
    assert.equal(readPreferences().theme, theme);
  }
  for (const theme of ['Dark', 'auto', '', null, 1, {}]) {
    global.window = { localStorage: storage({ 'injoffice.preferences.v1': JSON.stringify({ version: 1, defaultZoom: 150, showNavigation: false, theme }) }) };
    assert.deepEqual(readPreferences(), defaultPreferences, `theme ${JSON.stringify(theme)} rejects the whole blob`);
  }
});
