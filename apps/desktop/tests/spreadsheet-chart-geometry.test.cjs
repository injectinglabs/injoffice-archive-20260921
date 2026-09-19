const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

async function loadGeometry() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/spreadsheet-chart-geometry.ts'),
    platform: 'node',
    external: id => id === '@injoffice/xlsx-wasm' || id === '@injoffice/sheets/browser',
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

test('chart geometry refuses empty ranges, finite-normalizes values, and formats ticks', async () => {
  const { chartSelectionReason, chartGeometry, chartNumber, chartAnchor } = await loadGeometry();
  assert.match(chartSelectionReason({ row: 0, column: 0, end_row: 0, end_column: 1 }), /header row/);
  assert.equal(chartSelectionReason({ row: 0, column: 0, end_row: 2, end_column: 2 }), undefined);
  const chart = { editable: true, categories: ['A', 'B'], series: [{ values: [0, 10] }, { values: [5, 15] }] };
  const geometry = chartGeometry(chart);
  assert.equal(geometry.zero, 0);
  assert.equal(geometry.position(1), 1);
  assert.equal(geometry.ticks.length, 5);
  assert.throws(() => chartGeometry({ editable: true, categories: ['A'], series: [{ values: [Number.POSITIVE_INFINITY] }] }), /finite/);
  assert.equal(chartAnchor({ row: 1, column: 1, end_row: 3, end_column: 2 }).from_column, 4);
  assert.match(chartNumber(1234567890), /E|e/);
});
