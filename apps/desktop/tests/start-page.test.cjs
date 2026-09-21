const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const React = require('react');
const { create, act } = require('react-test-renderer');
global.IS_REACT_ACT_ENVIRONMENT = true;

async function loadStartPage() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/StartPage.tsx'),
    platform: 'node',
    external: id => /^react(?:\/|$)/.test(id),
    transform: { jsx: { runtime: 'automatic' } },
    plugins: [{
      name: 'assets',
      resolveId(id) { if (/\.(css|png)$/.test(id)) return '\0asset'; },
      load(id) { if (id === '\0asset') return 'export default ""'; },
    }],
  });
  try {
    const { output } = await bundle.generate({ format: 'cjs', codeSplitting: false });
    const mod = { exports: {} };
    new Function('require', 'module', 'exports', output[0].code)(require, mod, mod.exports);
    return mod.exports.default ?? mod.exports;
  } finally {
    await bundle.close();
  }
}

function recents() {
  return [
    { id: '1', name: 'Letter.docx', path: '/tmp/Letter.docx', updatedAt: 2 },
    { id: '2', name: 'Budget.xlsx', path: '/tmp/Budget.xlsx', updatedAt: 1 },
  ];
}

test('start page creates a document, filters recents, and disables Open while busy', async () => {
  const StartPage = await loadStartPage();
  const created = [];
  let view;
  await act(async () => {
    view = create(React.createElement(StartPage, {
      recentFiles: recents(),
      busy: false,
      available: true,
      onOpen() {},
      onCreate: format => created.push(format),
      onOpenRecent() {},
      onRemoveRecent() {},
    }));
  });
  const createDocument = view.root.findByProps({ 'aria-label': 'Create blank document' });
  await act(async () => createDocument.props.onClick());
  assert.deepEqual(created, ['docx']);
  const search = view.root.findByProps({ 'aria-label': 'Search recent files' });
  await act(async () => search.props.onChange({ target: { value: 'budget' } }));
  assert.equal(view.root.findAllByProps({ 'aria-label': 'Open Budget.xlsx' }).length, 1);
  assert.equal(view.root.findAllByProps({ 'aria-label': 'Open Letter.docx' }).length, 0);

  await act(async () => {
    view.update(React.createElement(StartPage, {
      recentFiles: recents(),
      busy: true,
      available: true,
      onOpen() {},
      onCreate: format => created.push(format),
      onOpenRecent() {},
      onRemoveRecent() {},
    }));
  });
  assert.equal(view.root.findByProps({ className: 'start-open-primary' }).props.disabled, true);
});

test('Home carries a Document Recovery card, every open document and folder names for recents', async () => {
  const { rolldown } = await import('rolldown');
  const StartPage = await loadStartPage();
  const recovered = [], discarded = [], selected = [];
  const now = Date.now();
  let view;
  await act(async () => {
    view = create(React.createElement(StartPage, {
      recentFiles: [{ id: '1', name: 'Letter.docx', path: '/Users/nick/Documents/Quarter 4/Letter.docx', updatedAt: now }],
      recoveries: [{ id: 'r1', name: 'Quarterly report.docx', updatedAt: now - 2 * 60 * 1000 }],
      onRecover: id => recovered.push(id),
      onDiscardRecovery: id => discarded.push(id),
      openDocuments: [{ key: 1, name: 'Quarterly report.docx' }, { key: 2, name: 'Budget.xlsx' }],
      onSelectDocument: key => selected.push(key),
      busy: false, available: true, onOpen() {}, onUpdates() {}, onCreate() {}, onOpenRecent() {}, onRemoveRecent() {},
    }));
  });
  // Recovery is a card inside the page with Office's title and a relative timestamp.
  const card = view.root.findByProps({ className: 'start-recovery' });
  assert.equal(card.findByProps({ id: 'start-recovery-title' }).children.join(''), 'Document Recovery');
  assert.match(card.findByProps({ className: 'start-recovery-file' }).findByType('small').children.join(''), /^Recovered · 2 min ago$/);
  await act(async () => card.findByProps({ className: 'start-recovery-open' }).props.onClick());
  await act(async () => card.findAllByType('button').find(button => button.props.children === 'Discard').props.onClick());
  assert.deepEqual([recovered, discarded], [['r1'], ['r1']]);
  // Every open document is listed, not just the active one.
  const current = view.root.findByProps({ className: 'start-current' });
  assert.deepEqual(current.findAllByType('button').map(button => button.props.title), ['Quarterly report.docx', 'Budget.xlsx']);
  await act(async () => current.findAllByType('button')[1].props.onClick());
  assert.deepEqual(selected, [2]);
  // Recent rows show the folder, with the full path in the tooltip.
  const row = view.root.findByProps({ 'aria-label': 'Open Letter.docx' });
  assert.equal(row.props.title, '/Users/nick/Documents/Quarter 4/Letter.docx');
  assert.equal(row.findByProps({ className: 'start-file-description' }).findByType('small').children.join(''), 'Quarter 4');
  // The rail's App updates entry carries an icon like its siblings.
  const updates = view.root.findAllByType('button').find(button => button.props.children?.some?.(child => child === 'App updates'));
  assert.ok(updates.findAllByType('svg').length, 'App updates has a glyph');
  await act(async () => view.unmount());

  // Relative times and folder labels are pure helpers.
  const bundle = await rolldown({ input: path.resolve(__dirname, '../src/StartPage.tsx'), platform: 'node', external: id => /^react(?:\/|$)/.test(id), transform: { jsx: { runtime: 'automatic' } }, plugins: [{ name: 'assets', resolveId(id) { if (/\.(css|png)$/.test(id)) return '\0asset'; }, load(id) { if (id === '\0asset') return 'export default ""'; } }] });
  const { output } = await bundle.generate({ format: 'cjs', codeSplitting: false });
  const mod = { exports: {} };
  new Function('require', 'module', 'exports', output[0].code)(require, mod, mod.exports);
  await bundle.close();
  const { relativeTime, folderLabel } = mod.exports;
  const base = 1_700_000_000_000;
  assert.equal(relativeTime(base, base + 5_000), 'Just now');
  assert.equal(relativeTime(base, base + 120_000), '2 min ago');
  assert.equal(relativeTime(base, base + 3_600_000), '1 hour ago');
  assert.equal(relativeTime(base, base + 7_200_000), '2 hours ago');
  assert.equal(relativeTime(base, base + 86_400_000), '1 day ago');
  assert.equal(folderLabel('/Users/nick/Documents/Quarter 4/Letter.docx'), 'Quarter 4');
  assert.equal(folderLabel('C:\\Users\\nick\\Reports\\Letter.docx'), 'Reports');
  assert.equal(folderLabel('Letter.docx'), 'Letter.docx');
});
