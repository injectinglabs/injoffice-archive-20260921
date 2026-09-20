const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const React = require('react');
const { create, act } = require('react-test-renderer');
global.IS_REACT_ACT_ENVIRONMENT = true;

async function load(file) {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src', file),
    platform: 'node',
    external: id => /^react(?:\/|$)/.test(id) || id.includes('packages/docs/src/'),
    transform: { jsx: { runtime: 'automatic' } },
    plugins: [{
      name: 'css',
      resolveId(id) { if (id.endsWith('.css')) return '\0css'; },
      load(id) { if (id === '\0css') return 'export default ""'; },
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

test('formatting toolbar reports bold and hyperlink apply as explicit patches', async () => {
  const FormattingToolbar = await load('FormattingToolbar.tsx');
  const HyperlinkControl = await load('HyperlinkControl.tsx');
  const patches = [];
  let view;
  await act(async () => {
    view = create(React.createElement(FormattingToolbar, {
      kind: 'docx',
      disabled: false,
      values: { font: 'Calibri', size: 11, bold: false },
      onChange: patch => patches.push(patch),
    }));
  });
  await act(async () => view.root.findByProps({ 'aria-label': 'Bold' }).props.onClick());
  assert.deepEqual(patches, [{ bold: true }]);

  const links = [];
  await act(async () => {
    view.update(React.createElement(HyperlinkControl, {
      disabled: false,
      onChange: url => links.push(url),
    }));
  });
  await act(async () => view.root.findAllByType('button')[0].props.onClick());
  const address = view.root.findByProps({ 'aria-label': 'Link address' });
  await act(async () => address.props.onChange({ target: { value: 'https://example.com' } }));
  await act(async () => view.root.findByProps({ 'aria-label': 'Link settings' }).props.onSubmit({ preventDefault() {} }));
  assert.deepEqual(links, ['https://example.com']);
});
