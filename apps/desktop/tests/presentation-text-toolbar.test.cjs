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
    external: id => /^react(?:\/|$)/.test(id) || id === '@injoffice/pptx-wasm',
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

const run = { text: 'Title', bold: false, italic: true, fontSizeHundredthPt: 2400, color: '20242B', fontFamily: 'Arial' };

test('presentation text toolbar maps the shared patch onto exact native run and paragraph fields', async () => {
  const { presentationTextPatch, presentationTextValues } = await load('PresentationTextToolbar.tsx');
  assert.deepEqual(presentationTextValues(run, 'center'), { font: 'Arial', size: 24, bold: false, italic: true, color: '#20242B', alignment: 'center' });
  assert.equal(presentationTextValues({ ...run, fontSizeHundredthPt: NaN }, 'left').size, undefined);
  assert.equal(presentationTextValues(undefined, 'left'), undefined);
  assert.deepEqual(presentationTextPatch({ bold: true }), { run: { bold: true } });
  assert.deepEqual(presentationTextPatch({ size: 18.5, color: '#b33c3c', font: 'Georgia' }), { run: { fontSizeHundredthPt: 1850, color: 'B33C3C', fontFamily: 'Georgia' } });
  assert.deepEqual(presentationTextPatch({ alignment: 'right' }), { align: 'right' });
  // Values outside the native subset never reach the transaction.
  assert.deepEqual(presentationTextPatch({ alignment: 'both' }), {});
  assert.deepEqual(presentationTextPatch({ size: 0, color: 'red' }), {});
});

test('presentation text toolbar drives run and alignment callbacks and keeps unsupported controls disabled with the reason', async () => {
  const { default: PresentationTextToolbar, PRESENTATION_TEXT_UNSUPPORTED } = await load('PresentationTextToolbar.tsx');
  const runs = [], aligns = [];
  let view;
  await act(async () => { view = create(React.createElement(PresentationTextToolbar, { run, align: 'left', disabled: false, onRunChange: patch => runs.push(patch), onAlignChange: align => aligns.push(align) })); });
  const bold = view.root.findByProps({ 'aria-label': 'Bold' });
  assert.equal(bold.props['aria-pressed'], false);
  assert.equal(view.root.findByProps({ 'aria-label': 'Italic' }).props['aria-pressed'], true);
  await act(async () => bold.props.onClick());
  await act(async () => view.root.findByProps({ 'aria-label': 'Font size' }).props.onChange({ target: { value: '36' } }));
  // Colour is an icon button with a palette, not a select: open it and pick the swatch.
  await act(async () => view.root.findByProps({ 'aria-label': 'Text color' }).props.onClick());
  await act(async () => view.root.findByProps({ 'aria-label': 'Blue' }).props.onClick());
  await act(async () => view.root.findByProps({ 'aria-label': 'Center' }).props.onClick());
  assert.deepEqual(runs, [{ bold: true }, { fontSizeHundredthPt: 3600 }, { color: '2459AD' }]);
  assert.deepEqual(aligns, ['center']);
  for (const [label, reason] of [['Underline', PRESENTATION_TEXT_UNSUPPORTED.underline], ['Bullets', PRESENTATION_TEXT_UNSUPPORTED.bullets]]) {
    const control = view.root.findByProps({ 'aria-label': label });
    assert.equal(control.props.disabled, true);
    assert.match(reason, /not supported by the native PPTX transaction/);
    assert.equal(control.props.title, reason);
  }
  // PowerPoint's alignment toggles only: the justify and distribute buttons Word has are not offered.
  const alignmentLabels = view.root.findByProps({ className: 'formatting-alignment' }).findAllByType('button').map(button => button.props['aria-label']);
  assert.deepEqual(alignmentLabels, ['Align left', 'Center', 'Align right']);
  assert.deepEqual(view.root.findAllByProps({ className: 'formatting-alignment' })[0].findAllByType('button').map(button => button.props['aria-pressed']), [true, false, false], 'the alignment in effect is the pressed toggle');

  await act(async () => { view.update(React.createElement(PresentationTextToolbar, { disabled: false, onRunChange: () => {}, onAlignChange: () => {} })); });
  assert.equal(view.root.findByProps({ 'aria-label': 'Bold' }).props.disabled, true);
  assert.equal(view.root.findByProps({ 'aria-label': 'Font family' }).props.disabled, true);
});
