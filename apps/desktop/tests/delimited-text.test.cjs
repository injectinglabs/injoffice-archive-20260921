const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

async function loadDelimited() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/delimitedText.ts'),
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

const utf8 = text => new TextEncoder().encode(text);

test('CSV/TSV parse keeps literal text, quoted commas, and rejects invalid UTF-8', async () => {
  const { parseDelimited, encodeDelimited } = await loadDelimited();
  assert.deepEqual(parseDelimited(utf8('a,b\r\n"c,d",e\n'), 'csv'), [['a', 'b'], ['c,d', 'e']]);
  assert.deepEqual(parseDelimited(utf8('a\tb\n'), 'tsv'), [['a', 'b']]);
  assert.throws(() => parseDelimited(utf8('a,b'), 'json'), /CSV or TSV/);
  assert.throws(() => parseDelimited(Buffer.from([0xff, 0xfe]), 'csv'), /UTF-8/);
  assert.throws(() => parseDelimited(utf8('a\0b'), 'csv'), /NUL/);
  const round = parseDelimited(encodeDelimited([['a,b', 'c"d'], ['1', '2']], 'csv'), 'csv');
  assert.deepEqual(round, [['a,b', 'c"d'], ['1', '2']]);
});
