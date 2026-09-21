const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
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

test('ribbon controls show the value in effect and stay empty for a mixed selection, never a placeholder word', async () => {
  const FormattingToolbar = await load('FormattingToolbar.tsx');
  // Placeholder wording is gone from every control the ribbon and the mini toolbar render.
  for (const file of ['FormattingToolbar.tsx', 'ParagraphToolbar.tsx', 'SelectionToolbar.tsx', 'Ribbon.tsx']) {
    const source = fs.readFileSync(path.resolve(__dirname, '../src', file), 'utf8');
    for (const banned of [/Inherited \w/, /From style/, /Mixed \/ /, /General alignment/, /placeholder="Inherited"/]) {
      assert.equal(banned.test(source), false, `${file} still shows "${banned}"`);
    }
  }
  const patches = [];
  let view;
  // A resolved caret: the font, size, colour and alignment in effect are visible in the controls.
  await act(async () => {
    view = create(React.createElement(FormattingToolbar, { kind: 'docx', disabled: false, onChange: patch => patches.push(patch), values: { font: 'Georgia', size: 14, bold: false, italic: false, underline: false, color: '#2459AD', alignment: 'center', characterEditable: true } }));
  });
  assert.equal(view.root.findByProps({ 'aria-label': 'Font family' }).props.value, 'Georgia');
  assert.equal(view.root.findByProps({ 'aria-label': 'Font size' }).props.value, '14');
  const pressed = () => view.root.findByProps({ className: 'formatting-alignment' }).findAllByType('button').map(button => button.props['aria-pressed']);
  assert.deepEqual(pressed(), [false, true, false, false, false], 'Word offers left/center/right/justify/distribute');
  // The colour button carries the colour as an underline; the palette applies one and closes.
  const bar = view.root.findByProps({ className: 'ribbon-color-bar' });
  assert.deepEqual(bar.props.style, { background: '#2459AD' });
  await act(async () => view.root.findByProps({ 'aria-label': 'Text color' }).props.onClick());
  assert.equal(view.root.findByProps({ 'aria-label': 'Blue' }).props['aria-pressed'], true, 'the current colour is the pressed swatch');
  await act(async () => view.root.findByProps({ 'aria-label': 'Red' }).props.onClick());
  assert.deepEqual(patches.at(-1), { color: '#B33C3C' });
  assert.equal(view.root.findAllByProps({ 'aria-label': 'Red' }).length, 0, 'picking a colour closes the palette');

  // A mixed selection: empty boxes with just the chevron, no pressed alignment, no colour underline.
  await act(async () => { view.update(React.createElement(FormattingToolbar, { kind: 'docx', disabled: false, onChange: patch => patches.push(patch), values: { characterEditable: true } })); });
  const empty = label => {
    const select = view.root.findByProps({ 'aria-label': label });
    assert.equal(select.props.value, '', `${label} is empty for a mixed selection`);
    assert.equal(select.props.children[0].props.value, '', `${label} offers a blank option`);
    assert.equal(select.props.children[0].props.children, undefined, `${label}'s blank option carries no text`);
  };
  empty('Font family');
  empty('Font size');
  assert.deepEqual(pressed(), [false, false, false, false, false]);
  assert.equal(view.root.findByProps({ className: 'ribbon-color-bar' }).props.style, undefined);
  await act(async () => view.unmount());
});

test('paragraph toolbar can create lists when the document has no numbering definitions', async () => {
  const ParagraphToolbar = await load('ParagraphToolbar.tsx'); const patches = []; let view;
  await act(async () => { view = create(React.createElement(ParagraphToolbar, { properties: {}, styles: [], numbering: [], disabled: false, onChange: patch => patches.push(patch) })); });
  try {
    const list = view.root.findByProps({ 'aria-label': 'Paragraph list' });
    assert.equal(list.props.disabled, false);
    for (const value of ['@bullet', '@decimal', '0']) await act(async () => list.props.onChange({ target: { value } }));
    assert.deepEqual(patches, [{ numbering_kind: 'bullet', numbering_level: 0 }, { numbering_kind: 'decimal', numbering_level: 0 }, { numbering_num_id: '0', numbering_level: 0 }]);
  } finally { await act(async () => view.unmount()); }
});

test('unsupported paragraph-style writes stay read-only even when style metadata is present', async () => {
  const ParagraphToolbar = await load('ParagraphToolbar.tsx'); let view;
  await act(async () => { view = create(React.createElement(ParagraphToolbar, { properties: { paragraph_style_id: 'Heading1' }, styles: [{ id: 'Heading1', name: 'Heading 1' }], disabled: false, onChange() { throw new Error('unsupported style write'); } })); });
  try {
    const style = view.root.findByProps({ 'aria-label': 'Paragraph style' });
    assert.equal(style.props.disabled, true);
    assert.equal(style.props.onChange, undefined);
    assert.match(style.props.title, /not supported/);
    assert.equal(style.props.value, 'Heading1');
  } finally { await act(async () => view.unmount()); }
});
