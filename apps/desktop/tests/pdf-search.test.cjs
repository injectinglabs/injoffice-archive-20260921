const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

async function loadSearch() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/pdf-search.ts'),
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

test('PDF search matches across parts, reports pages, and honors abort', async () => {
  const { findPdfTextMatches, searchPdfDocument } = await loadSearch();
  const hits = findPdfTextMatches(['Hello', 'world'], 'lo wo');
  assert.equal(hits.length, 1);
  assert.deepEqual(hits[0].spans, [0, 1]);
  const pages = { 1: ['alpha'], 2: ['beta alpha'], 3: ['gamma'] };
  const scanned = [];
  const result = await searchPdfDocument(3, async page => pages[page], 'alpha', new AbortController().signal, page => scanned.push(page));
  assert.deepEqual(result.hits, [{ page: 1, match: 0 }, { page: 2, match: 0 }]);
  assert.deepEqual(scanned, [1, 2, 3]);
  assert.equal(result.limited, false);
  const controller = new AbortController();
  controller.abort();
  await assert.rejects(searchPdfDocument(3, async () => ['x'], 'x', controller.signal), /aborted/i);
  await assert.rejects(searchPdfDocument(0, async () => ['x'], 'x', new AbortController().signal), /Enter text/);
});
