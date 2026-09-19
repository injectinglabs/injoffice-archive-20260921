const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

async function loadDelimited() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/spreadsheetDelimited.ts'),
    platform: 'node',
    external: id => id === '@injoffice/sheets/browser' || id === '@injoffice/xlsx-wasm' || id.includes('packages/formulas/'),
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

function workbook() {
  return {
    revision: 'rev-1',
    unsupported: [],
    sheets: [{
      id: '1',
      cells: [
        { row: 0, column: 0, value: { kind: 'string', text: 'a,b' } },
        { row: 0, column: 1, formula: { type: 'normal', text: 'A1' }, ref: 'B1' },
      ],
    }],
  };
}

test('exportDelimitedSheet writes literal values or formula source and refuses stale caches', async () => {
  const { exportDelimitedSheet } = await loadDelimited();
  const { parseDelimited } = await (async () => {
    const { rolldown } = await import('rolldown');
    const bundle = await rolldown({ input: path.resolve(__dirname, '../src/delimitedText.ts'), platform: 'node' });
    try {
      const { output } = await bundle.generate({ format: 'cjs', codeSplitting: false });
      const mod = { exports: {} };
      new Function('require', 'module', 'exports', output[0].code)(require, mod, mod.exports);
      return mod.exports;
    } finally { await bundle.close(); }
  })();
  const formulas = parseDelimited(exportDelimitedSheet(workbook(), '1', 'csv', 'formulas'), 'csv');
  assert.deepEqual(formulas, [['a,b', '=A1']]);
  assert.throws(() => exportDelimitedSheet(workbook(), '1', 'csv', 'values'), /Recalculate/);
  const calculated = {
    revision: 'rev-1',
    result: { cells: [{ sheetId: '1', row: 0, column: 1, formula: '=A1', status: 'calculated', value: { kind: 'string', value: 'a,b' } }] },
  };
  const values = parseDelimited(exportDelimitedSheet(workbook(), '1', 'csv', 'values', calculated), 'csv');
  assert.deepEqual(values, [['a,b', 'a,b']]);
});
