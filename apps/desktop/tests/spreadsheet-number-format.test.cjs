const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const React = require('react');
const { create, act } = require('react-test-renderer');
global.IS_REACT_ACT_ENVIRONMENT = true;

async function load() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/SpreadsheetNumberFormat.tsx'),
    platform: 'node',
    external: id => /^react(?:\/|$)/.test(id),
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

test('number format helpers stay inside the qualified native display subset', async () => {
  const { parseNumericFormat, adjustDecimals, NUMBER_FORMAT_PRESETS, MAX_DECIMAL_PLACES } = await load();
  assert.deepEqual(parseNumericFormat('General'), { prefix: '', grouped: false, decimals: 0, percent: false, suffix: '' });
  assert.deepEqual(parseNumericFormat('"$"#,##0.00'), { prefix: '$', grouped: true, decimals: 2, percent: false, suffix: '' });
  assert.equal(parseNumericFormat('yyyy-mm-dd'), undefined);
  assert.equal(adjustDecimals('General', 1), '0.0');
  assert.equal(adjustDecimals(undefined, -1), undefined);
  assert.equal(adjustDecimals('0%', 1), '0.0%');
  assert.equal(adjustDecimals('0.0%', -1), '0%');
  assert.equal(adjustDecimals('"$"#,##0.00', 1), '"$"#,##0.000');
  assert.equal(adjustDecimals('#,##0.00', -1), '#,##0.0');
  assert.equal(adjustDecimals(`0.${'0'.repeat(MAX_DECIMAL_PLACES)}`, 1), undefined);
  assert.equal(adjustDecimals('yyyy-mm-dd', 1), undefined);
  // Every preset and every adjusted preset parses back, so the grid renders it instead of "Stored value shown".
  for (const preset of Object.values(NUMBER_FORMAT_PRESETS)) if (preset !== 'yyyy-mm-dd') {
    assert.ok(parseNumericFormat(preset), preset);
    assert.ok(parseNumericFormat(adjustDecimals(preset, 1)), `${preset} +1`);
  }
});

test('number format control drives the number_format command path and borders stay fail-closed', async () => {
  const { default: SpreadsheetNumberFormat, SpreadsheetBorders, BORDERS_UNSUPPORTED, DECIMALS_UNSUPPORTED } = await load();
  const changes = [];
  let view;
  await act(async () => { view = create(React.createElement(SpreadsheetNumberFormat, { numberFormat: 'General', disabled: false, onChange: value => changes.push(value) })); });
  const by = label => view.root.findByProps({ 'aria-label': label });
  assert.equal(by('Cell number format').props.value, 'General');
  assert.equal(by('Decrease decimals').props.disabled, true);
  assert.equal(by('Decrease decimals').props.title, DECIMALS_UNSUPPORTED);
  await act(async () => by('Percent format').props.onClick());
  await act(async () => by('Currency format').props.onClick());
  await act(async () => by('Comma style').props.onClick());
  await act(async () => by('Increase decimals').props.onClick());
  await act(async () => by('Cell number format').props.onChange({ target: { value: '0.00' } }));
  assert.deepEqual(changes, ['0%', '"$"#,##0.00', '#,##0.00', '0.0', '0.00']);

  await act(async () => { view.update(React.createElement(SpreadsheetNumberFormat, { numberFormat: '0.00%', disabled: false, onChange: value => changes.push(value) })); });
  assert.equal(by('Percent format').props['aria-pressed'], true);
  assert.equal(by('Currency format').props['aria-pressed'], false);
  await act(async () => by('Decrease decimals').props.onClick());
  assert.equal(changes.at(-1), '0.0%');

  await act(async () => { view.update(React.createElement(SpreadsheetNumberFormat, { numberFormat: 'yyyy-mm-dd', disabled: false, onChange: () => {} })); });
  assert.equal(by('Increase decimals').props.disabled, true);
  assert.equal(by('Increase decimals').props.title, DECIMALS_UNSUPPORTED);
  assert.ok(by('Cell number format').props.children.some(option => option.props.value === 'yyyy-mm-dd'));

  await act(async () => { view.update(React.createElement(SpreadsheetNumberFormat, { numberFormat: '[$-409]#,##0', disabled: true, onChange: () => {} })); });
  assert.equal(by('Percent format').props.disabled, true);
  assert.ok(by('Cell number format').props.children.some(option => option.props.children === 'Existing: [$-409]#,##0'));

  await act(async () => { view.update(React.createElement(SpreadsheetBorders)); });
  const borders = view.root.findByProps({ 'aria-label': 'Cell borders' });
  assert.equal(borders.props.disabled, true);
  assert.equal(borders.props.title, BORDERS_UNSUPPORTED);
  assert.match(BORDERS_UNSUPPORTED, /not supported by the native XLSX transaction/);
});
