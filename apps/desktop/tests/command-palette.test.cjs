const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const React = require('react');
const { create, act } = require('react-test-renderer');
global.IS_REACT_ACT_ENVIRONMENT = true;

async function loadPalette() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/CommandPalette.tsx'),
    platform: 'node',
    external: id => /^react(?:\/|$)/.test(id),
    transform: { jsx: { runtime: 'automatic' } },
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

test('command palette filters and runs the matching command', async () => {
  const CommandPalette = await loadPalette();
  const ran = [];
  let closed = 0;
  let view;
  await act(async () => {
    view = create(React.createElement(CommandPalette, {
      onClose: () => { closed += 1; },
      commands: [
        { id: 'new', label: 'New document', run: () => ran.push('new') },
        { id: 'save', label: 'Save', detail: 'Current file', run: () => ran.push('save') },
      ],
    }));
  });
  const search = view.root.findByProps({ 'aria-label': 'Search workspace commands' });
  await act(async () => search.props.onChange({ target: { value: 'save' } }));
  assert.equal(view.root.findAllByProps({ role: 'option' }).length, 1);
  await act(async () => search.props.onKeyDown({ key: 'Enter', preventDefault() {} }));
  assert.deepEqual(ran, ['save']);
  assert.equal(closed, 1);
});


test('command palette opens modally, closes on Escape/cancel and backdrop, and keeps interior clicks open', async () => {
  const CommandPalette = await loadPalette();
  let modal = 0, closed = 0, prevented = 0, view;
  const dialog = { showModal() { modal++; }, querySelector() {}, getBoundingClientRect: () => ({ left: 100, right: 500, top: 100, bottom: 400 }) };
  await act(async () => { view = create(React.createElement(CommandPalette, { commands: [], onClose: () => closed++ }), { createNodeMock: node => node.type === 'dialog' ? dialog : null }); });
  assert.equal(modal, 1, 'showModal enables native Escape cancellation');
  const element = view.root.findByType('dialog');
  assert.equal(element.props.open, undefined, 'an open attribute would prevent showModal from making it modal');
  element.props.onCancel({ preventDefault() { prevented++; } });
  assert.equal(closed, 1); assert.equal(prevented, 1);
  element.props.onClick({ target: dialog, clientX: 150, clientY: 150 });
  element.props.onClick({ target: {}, clientX: 0, clientY: 0 });
  assert.equal(closed, 1);
  element.props.onClick({ target: dialog, clientX: 50, clientY: 50 });
  assert.equal(closed, 2);
  const search = view.root.findByProps({ 'aria-label': 'Search workspace commands' });
  search.props.onKeyDown({ key: 'Escape', preventDefault() {}, stopPropagation() {} });
  assert.equal(closed, 3, 'Escape from the search field closes the palette');
  element.props.onKeyDown({ key: 'Escape', preventDefault() {}, stopPropagation() {} });
  assert.equal(closed, 4, 'Escape closes the palette when focus is on the dialog');
  await act(async () => view.unmount());
});
