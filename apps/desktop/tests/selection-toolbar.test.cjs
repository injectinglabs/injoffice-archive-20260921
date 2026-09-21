const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const React = require('react');
const { create, act } = require('react-test-renderer');
global.IS_REACT_ACT_ENVIRONMENT = true;

async function load() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/SelectionToolbar.tsx'),
    platform: 'node',
    external: id => /^react(?:\/|$)/.test(id),
    transform: { jsx: { runtime: 'automatic' } },
    plugins: [{ name: 'css', resolveId(id) { if (id.endsWith('.css')) return '\0css'; }, load(id) { if (id === '\0css') return 'export default ""'; } }],
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

test('mini toolbar renders the ribbon FormattingToolbar above the selection and drives the same patch path', async () => {
  assert.equal(fs.existsSync(path.resolve(__dirname, '../src/selection-toolbar.css')), true);
  const { default: SelectionToolbar } = await load();
  const patches = [];
  let view;
  await act(async () => {
    view = create(React.createElement(SelectionToolbar, { anchor: { left: 120, top: 300, bottom: 318, width: 80 }, disabled: false, values: { font: 'Calibri', size: 11, bold: false, italic: true, underline: false, color: '#2459AD', alignment: 'left', characterEditable: true }, onChange: patch => patches.push(patch) }));
  });
  const toolbar = view.root.findByProps({ role: 'toolbar' });
  assert.equal(toolbar.props['aria-label'], 'Selection formatting');
  assert.equal(toolbar.props['data-placement'], 'above');
  // Without a DOM the toolbar keeps the anchor coordinates and stays hidden until it can be measured.
  assert.equal(toolbar.props.style.left, 120);
  assert.equal(toolbar.props.style.top, 300);
  assert.equal(toolbar.props.style.visibility, 'hidden');
  const inner = view.root.findByProps({ 'aria-label': 'Formatting' });
  assert.equal(inner.props.className, 'formatting-toolbar');
  for (const label of ['Font family', 'Font size', 'Bold', 'Italic', 'Underline', 'Text color', 'Align left', 'Center', 'Align right']) assert.equal(view.root.findAllByProps({ 'aria-label': label }).length, 1, label);
  await act(async () => view.root.findByProps({ 'aria-label': 'Bold' }).props.onClick());
  await act(async () => view.root.findByProps({ 'aria-label': 'Italic' }).props.onClick());
  await act(async () => view.root.findByProps({ 'aria-label': 'Font size' }).props.onChange({ target: { value: '14' } }));
  await act(async () => view.root.findByProps({ 'aria-label': 'Center' }).props.onClick());
  assert.deepEqual(patches, [{ bold: true }, { italic: false }, { size: 14 }, { alignment: 'center' }]);
  assert.equal(view.root.findByProps({ 'aria-label': 'Italic' }).props['aria-pressed'], true);
  // Busy editors disable every control, exactly like the ribbon.
  await act(async () => view.update(React.createElement(SelectionToolbar, { anchor: { left: 120, top: 300, bottom: 318, width: 80 }, disabled: true, values: { bold: false }, onChange: patch => patches.push(patch) })));
  assert.equal(view.root.findByProps({ 'aria-label': 'Bold' }).props.disabled, true);
  await act(async () => view.unmount());
});

test('mini toolbar renders nothing without a selection or a DOM', async () => {
  const { default: SelectionToolbar } = await load();
  let view;
  await act(async () => { view = create(React.createElement(SelectionToolbar, { disabled: false, onChange() {} })); });
  assert.equal(view.toJSON(), null);
  await act(async () => view.unmount());
});

test('selection geometry: only non-empty selections inside the canvas anchor the toolbar, which flips below when clipped', async () => {
  const { selectionAnchorRect, toolbarPosition } = await load();
  const rect = { left: 100, top: 400, bottom: 418, width: 60, height: 18 };
  const inside = { closest: selector => selector === '.office-document-canvas' ? {} : null };
  const outside = { closest: () => null };
  const selection = (text, element, collapsed = false) => ({ rangeCount: 1, isCollapsed: collapsed, toString: () => text, getRangeAt: () => ({ commonAncestorContainer: { nodeType: 3, parentElement: element }, getClientRects: () => [rect], getBoundingClientRect: () => rect }) });
  assert.deepEqual(selectionAnchorRect(selection('menu', inside), '.office-document-canvas'), { left: 100, top: 400, bottom: 418, width: 60 });
  assert.equal(selectionAnchorRect(selection('menu', outside), '.office-document-canvas'), undefined);
  assert.equal(selectionAnchorRect(selection('', inside), '.office-document-canvas'), undefined);
  assert.equal(selectionAnchorRect(selection('menu', inside, true), '.office-document-canvas'), undefined);
  assert.equal(selectionAnchorRect({ rangeCount: 0, isCollapsed: true, toString: () => '' }, '.office-document-canvas'), undefined);
  assert.equal(selectionAnchorRect(null, '.office-document-canvas'), undefined);
  const size = { width: 500, height: 40 }, viewport = { width: 1200, height: 800 };
  assert.deepEqual(toolbarPosition({ left: 300, top: 400, bottom: 418, width: 60 }, size, viewport), { x: 300, y: 352, placement: 'above' });
  assert.deepEqual(toolbarPosition({ left: 300, top: 20, bottom: 38, width: 60 }, size, viewport), { x: 300, y: 46, placement: 'below' });
  assert.deepEqual(toolbarPosition({ left: 1000, top: 400, bottom: 418, width: 60 }, size, viewport), { x: 692, y: 352, placement: 'above' });
  assert.deepEqual(toolbarPosition({ left: 2, top: 400, bottom: 418, width: 60 }, size, viewport), { x: 8, y: 352, placement: 'above' });
});

test('the document editor mounts the mini toolbar on the ribbon values and changeFormatting', () => {
  const office = fs.readFileSync(path.resolve(__dirname, '../src/OfficeEditor.tsx'), 'utf8');
  assert.match(office, /import SelectionToolbar from '\.\/SelectionToolbar'/);
  assert.match(office, /<SelectionToolbar values=\{toolbarValues\} disabled=\{busy\|\|composing\} onChange=\{patch=>void changeFormatting\(patch\)\} \/>/);
});
