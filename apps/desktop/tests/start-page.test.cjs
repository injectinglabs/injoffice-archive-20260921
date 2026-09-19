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
