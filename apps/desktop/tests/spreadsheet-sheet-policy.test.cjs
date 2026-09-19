const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

async function loadPolicy() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/spreadsheetSheetPolicy.ts'),
    platform: 'node',
    external: id => id === '@injoffice/sheets/browser',
  });
  try {
    const { output } = await bundle.generate({ format: 'cjs', codeSplitting: false });
    const mod = { exports: {} };
    const req = id => (id === '@injoffice/sheets/browser' ? { editableDefinedName: name => !String(name).startsWith('_') } : require(id));
    new Function('require', 'module', 'exports', output[0].code)(req, mod, mod.exports);
    return mod.exports;
  } finally {
    await bundle.close();
  }
}

function workbook(extra = {}) {
  return {
    unsupported: [],
    defined_names: [],
    sheets: [{ id: '1', editable: true, state: 'visible', cells: [], auto_filter: false }],
    ...extra,
  };
}

test('sheetLifecycleReason refuses delete of the last visible sheet and add on a read-only workbook', async () => {
  const { sheetLifecycleReason } = await loadPolicy();
  assert.equal(sheetLifecycleReason(workbook(), '1', 'delete'), 'Keep at least one visible worksheet.');
  const two = workbook({ sheets: [
    { id: '1', editable: true, state: 'visible', cells: [], auto_filter: false },
    { id: '2', editable: true, state: 'visible', cells: [], auto_filter: false },
  ] });
  assert.equal(sheetLifecycleReason(two, '1', 'delete'), undefined);
  assert.equal(sheetLifecycleReason(workbook({ sheets: [{ id: '1', editable: false, state: 'visible', cells: [], auto_filter: false }] }), '1', 'add'), 'This worksheet is read-only.');
  assert.match(sheetLifecycleReason(two, '1', 'delete') ?? '', /^$/);
  const formulas = workbook({ sheets: [
    { id: '1', editable: true, state: 'visible', cells: [{ formula: 'A1' }], auto_filter: false },
    { id: '2', editable: true, state: 'visible', cells: [], auto_filter: false },
  ] });
  assert.match(sheetLifecycleReason(formulas, '2', 'delete'), /formulas/);
});
