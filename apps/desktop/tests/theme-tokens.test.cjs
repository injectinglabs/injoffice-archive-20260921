// The app chrome is painted only through the theme tokens declared in styles.css,
// so a dark theme is a matter of the token blocks, never of per-rule overrides.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const src = path.resolve(__dirname, '../src');
const styles = fs.readFileSync(path.join(src, 'styles.css'), 'utf8');

function tokenBlock(selectorPattern) {
  const match = styles.match(new RegExp(`${selectorPattern}\\s*\\{([^}]*)\\}`));
  assert.ok(match, `styles.css declares a block for ${selectorPattern}`);
  const tokens = new Map();
  for (const [, name, value] of match[1].matchAll(/(--[\w-]+)\s*:\s*([^;]+);/g)) tokens.set(name, value.trim());
  return { body: match[1], tokens };
}

const themed = ['--surface', '--surface-raised', '--surface-muted', '--canvas', '--hover', '--border', '--border-subtle', '--border-strong', '--text', '--text-secondary', '--text-muted', '--text-faint', '--accent', '--accent-fill', '--accent-hover', '--accent-text', '--accent-soft', '--accent-border', '--selection', '--danger', '--danger-fill', '--danger-soft', '--danger-border', '--warning', '--warning-text', '--warning-soft', '--shadow', '--shadow-soft', '--backdrop', '--docx', '--docx-soft', '--xlsx', '--xlsx-fill', '--xlsx-hover', '--xlsx-soft', '--xlsx-border', '--xlsx-selection', '--pptx', '--pptx-soft', '--pdf', '--pdf-soft'];
const documentTokens = ['--document-surface', '--document-edit-surface', '--document-text', '--document-heading', '--document-muted', '--document-border', '--document-link', '--document-focus', '--document-selection'];

test('styles.css declares every chrome token in light, system-dark, and forced-dark blocks', () => {
  const light = tokenBlock(':root');
  const systemDark = tokenBlock('@media \\(prefers-color-scheme: dark\\)\\s*\\{\\s*:root:not\\(\\[data-theme="light"\\]\\)');
  const forcedDark = tokenBlock(':root\\[data-theme="dark"\\]');
  assert.match(light.body, /color-scheme:\s*light/);
  assert.match(systemDark.body, /color-scheme:\s*dark/);
  assert.match(forcedDark.body, /color-scheme:\s*dark/);
  for (const token of themed) {
    assert.ok(light.tokens.has(token), `${token} has a light value`);
    assert.ok(systemDark.tokens.has(token), `${token} has a system-dark value`);
    assert.ok(forcedDark.tokens.has(token), `${token} has a forced-dark value`);
    assert.equal(systemDark.tokens.get(token), forcedDark.tokens.get(token), `${token} is the same in both dark blocks`);
  }
  assert.deepEqual([...systemDark.tokens.keys()].sort(), [...forcedDark.tokens.keys()].sort());
  // Document surfaces never change with the theme, so the dark blocks must not touch them.
  for (const token of documentTokens) {
    assert.ok(light.tokens.has(token), `${token} is declared once on :root`);
    assert.ok(!systemDark.tokens.has(token) && !forcedDark.tokens.has(token), `${token} is not overridden in dark blocks`);
  }
  assert.equal(light.tokens.get('--document-surface'), '#fff');
  // The ribbon and other lanes use the legacy names; they must resolve to the tokens.
  assert.equal(light.tokens.get('--ink'), 'var(--text)');
  assert.equal(light.tokens.get('--muted'), 'var(--text-muted)');
  assert.equal(light.tokens.get('--line'), 'var(--border)');
  assert.equal(light.tokens.get('--blue'), 'var(--accent)');
});

test('chrome stylesheets carry no colour literals outside the token blocks and document-layer paint', () => {
  // Literals that intentionally remain: PDF annotation paint drawn on the page, and the paper
  // mock-ups inside the start page's create cards (document illustrations, not chrome).
  const documentLayer = new Set(['#fff9d7', '#fff9', '#fff0a7', '#ffed91', '#ffd52f50', '#ffcb3e66', '#ec922b88', '#c7ae55', '#be6b15', '#b99726', '#5b4a0f', '#4f4617', '#493e16', '#147cf360',
    '#e0e9e4', '#e0eee6', '#edf4ef', '#69a383', '#eef6f155', '#dcbfaf', '#e8dbd4', '#e1d5ce', '#ccdbd2', '#dce4f0', '#bbcee8']);
  const chrome = ['styles.css', 'start-page.css', 'office-editor.css', 'presentation-editor.css', 'spreadsheet.css', 'spreadsheet-charts.css', 'pdf-editor.css', 'open-error.css', 'updates-dialog.css'];
  for (const name of chrome) {
    let css = fs.readFileSync(path.join(src, name), 'utf8');
    if (name === 'styles.css') css = css.replace(/:root[^{]*\{[^}]*\}/g, '');
    const literals = [...css.matchAll(/#[0-9a-fA-F]{3,8}\b|(?<![\w-])(?:white|black)(?![\w-])/g)].map(m => m[0].toLowerCase()).filter(literal => !documentLayer.has(literal));
    assert.deepEqual(literals, [], `${name} paints chrome with literals: ${[...new Set(literals)].join(' ')}`);
  }
});

test('index.html lets the renderer follow both colour schemes', () => {
  const html = fs.readFileSync(path.resolve(__dirname, '../index.html'), 'utf8');
  assert.match(html, /<meta name="color-scheme" content="light dark" \/>/);
});
