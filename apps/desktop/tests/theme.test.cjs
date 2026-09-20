const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

async function loadTheme() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({ input: path.resolve(__dirname, '../src/theme.ts'), platform: 'node', transform: { jsx: { runtime: 'automatic' } } });
  try {
    const { output } = await bundle.generate({ format: 'cjs', codeSplitting: false });
    const mod = { exports: {} };
    new Function('require', 'module', 'exports', output[0].code)(require, mod, mod.exports);
    return mod.exports;
  } finally { await bundle.close(); }
}

test('applyTheme pins data-theme for a manual choice and removes it to follow the system', async () => {
  const { applyTheme } = await loadTheme();
  const root = { dataset: {} };
  applyTheme('dark', root);
  assert.equal(root.dataset.theme, 'dark', 'dark matches the :root[data-theme="dark"] block in styles.css');
  applyTheme('light', root);
  assert.equal(root.dataset.theme, 'light', 'light opts out of the prefers-color-scheme block');
  applyTheme('system', root);
  assert.equal('theme' in root.dataset, false, 'system leaves the OS setting in charge');
  applyTheme('dark', undefined);
});
