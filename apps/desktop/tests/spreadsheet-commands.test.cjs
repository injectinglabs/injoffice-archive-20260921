const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

async function loadCommands() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/spreadsheetCommands.ts'),
    platform: 'node',
    external: id => id === '@injoffice/sheets/browser' || id === '@injoffice/xlsx-wasm',
  });
  try {
    const { output } = await bundle.generate({ format: 'cjs', codeSplitting: false });
    const mod = { exports: {} };
    const req = id => (id === '@injoffice/sheets/browser' || id === '@injoffice/xlsx-wasm'
      ? { formatNativeSheetCellDisplayV2: () => { throw new Error('WASM display is mocked'); } }
      : require(id));
    new Function('require', 'module', 'exports', output[0].code)(req, mod, mod.exports);
    return mod.exports;
  } finally {
    await bundle.close();
  }
}

function workbook(extra = {}) {
  return {
    source: { package_sha256: 'rev-1' },
    styles: [],
    sheets: [{
      id: '1',
      editable: true,
      state: 'visible',
      merged_ranges: [],
      cells: [
        { row: 0, column: 0, value: { kind: 'string', text: 'a' } },
        { row: 0, column: 1, value: { kind: 'string', text: 'b' } },
        { row: 1, column: 1, value: { kind: 'number', lexical: '2' } },
      ],
    }],
    ...extra,
  };
}

test('parseSelection, visible-row windows, 2x2 copy, and fail-closed recovery drafts', async () => {
  const { parseSelection, visibleRowWindow, copySelection, validateRecoveryDraft } = await loadCommands();
  assert.deepEqual(parseSelection('A1:B2'), { anchor: { row: 0, column: 0 }, end: { row: 1, column: 1 } });
  assert.deepEqual(parseSelection('B12'), { anchor: { row: 11, column: 1 }, end: { row: 11, column: 1 } });
  assert.throws(() => parseSelection('A1:B2:C3'), /rectangular/);
  const hidden = new Set(Array.from({ length: 150 }, (_, i) => i + 1));
  assert.deepEqual(visibleRowWindow(0, 4, hidden, 200), [0, 151, 152, 153]);
  const sheet = workbook().sheets[0];
  assert.equal(copySelection(sheet, parseSelection('A1:B2')), 'a\tb\r\n\t2');
  const book = workbook();
  const draft = { version: 1, format: 'xlsx', revision: 'rev-1', sheetId: '1', row: 0, column: 0, value: '=A1' };
  assert.deepEqual(validateRecoveryDraft(book, draft), draft);
  assert.throws(() => validateRecoveryDraft(book, { ...draft, extra: true }), /match this workbook/);
});
