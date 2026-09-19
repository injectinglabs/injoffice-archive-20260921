const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

async function loadArrange() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/presentationArrange.ts'),
    platform: 'node',
    external: id => id === '@injoffice/pptx-native' || id === '@injoffice/pptx-wasm' || id.includes('playground/src/pptxRoundTrip'),
  });
  try {
    const { output } = await bundle.generate({ format: 'cjs', codeSplitting: false });
    const mod = { exports: {} };
    const req = id => (id.includes('pptxRoundTrip') ? { editablePptxTextTargets: () => [], editablePptxShapeTargets: () => [] } : require(id));
    new Function('require', 'module', 'exports', output[0].code)(req, mod, mod.exports);
    return mod.exports;
  } finally {
    await bundle.close();
  }
}

test('arrange selection toggles keys and refuses fewer than two objects', async () => {
  const { toggleArrangeSelection, arrangeCommand } = await loadArrange();
  assert.deepEqual(toggleArrangeSelection(['a'], 'b'), ['a', 'b']);
  assert.deepEqual(toggleArrangeSelection(['a', 'b'], 'a'), ['b']);
  const deck = { sourceRevision: 'rev', slides: [{ elements: [] }] };
  assert.throws(() => arrangeCommand(deck, 0, ['only'], 'left', 'op'), /Select 2/);
  assert.throws(() => arrangeCommand(deck, 0, ['a', 'b'], 'horizontal', 'op'), /Select 3/);
  assert.throws(() => arrangeCommand(deck, 0, ['a', 'a'], 'left', 'op'), /only once/);
});
