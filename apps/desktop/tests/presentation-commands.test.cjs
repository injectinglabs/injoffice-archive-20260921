const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

async function loadCommands() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/presentationCommands.ts'),
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

function deck(slides) {
  return { sourceRevision: 'rev', slides };
}

test('presentation commands key elements, refuse deleting the last slide, and project grouped coordinates', async () => {
  const { elementKey, structureCommand, positionElements } = await loadCommands();
  assert.equal(elementKey({ id: 'local', source: { partName: 'ppt/slides/slide1.xml', objectId: '3' } }), 'ppt/slides/slide1.xml\03');
  const two = deck([{ id: 's1' }, { id: 's2' }]);
  assert.equal(structureCommand(two, 0, 'duplicate', 'op').operations[0].kind, 'slide.duplicate');
  assert.deepEqual(structureCommand(two, 0, 'next', 'op').operations[0].slideIds, ['s2', 's1']);
  assert.throws(() => structureCommand(deck([{ id: 's1' }]), 0, 'delete', 'op'), /last slide/);
  const grouped = positionElements([{
    id: 'g', kind: 'group', compatibility: { status: 'editable' },
    transform: { x: 100, y: 0, cx: 200, cy: 100 },
    childTransform: { x: 0, y: 0, cx: 100, cy: 50 },
    children: [{ id: 'c', kind: 'shape', compatibility: { status: 'editable' }, transform: { x: 10, y: 5, cx: 20, cy: 10 } }],
  }]);
  assert.equal(grouped[0].grouped, true);
  assert.equal(grouped[0].rect.x, 120);
  assert.equal(grouped[0].rect.cx, 40);
});
