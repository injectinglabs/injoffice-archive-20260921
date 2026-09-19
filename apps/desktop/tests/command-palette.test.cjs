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
