const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

async function loadRange() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/document-range.ts'),
    platform: 'node',
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

test('paragraphTextOffset counts UTF-16 authored text and rejects out-of-range runs', async () => {
  const { paragraphTextOffset } = await loadRange();
  assert.equal(paragraphTextOffset([3, 5, 2], 1, 2), 5);
  assert.equal(paragraphTextOffset([3, 5, 2], 0, 0), 0);
  assert.equal(paragraphTextOffset([3, 5, 2], 2, 2), 10);
  assert.equal(paragraphTextOffset([3, 5, 2], 1, 6), undefined);
  assert.equal(paragraphTextOffset([3, 5, 2], -1, 0), undefined);
  assert.equal(paragraphTextOffset([3, 5, 2], 1, 1.5), undefined);
});
