const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

async function loadCaret() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/document-caret.ts'),
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

// One short first line, a full second line, and an empty last paragraph that fills its line.
function boxes() {
  return [
    { left: 100, right: 180, top: 100, bottom: 122 },
    { left: 100, right: 700, top: 130, bottom: 152 },
    { left: 100, right: 700, top: 160, bottom: 182 },
  ];
}

test('caretTargetAt has nothing to offer without an editable run', async () => {
  const { caretTargetAt } = await loadCaret();
  assert.equal(caretTargetAt([], 300, 300), undefined);
});

test('caretTargetAt keeps a click inside the run it hit', async () => {
  const { caretTargetAt } = await loadCaret();
  const hit = caretTargetAt(boxes(), 140, 110);
  assert.equal(hit.index, 0);
  assert.deepEqual([hit.x, hit.y], [140, 110]);
});

test('caretTargetAt stays on the clicked line past the end of its text', async () => {
  const { caretTargetAt } = await loadCaret();
  const hit = caretTargetAt(boxes(), 640, 110);
  assert.equal(hit.index, 0, 'the short first line, not the long second one');
  assert.equal(hit.x, 179.5, 'clamped to the last character of that line');
  assert.equal(hit.y, 110);
});

test('caretTargetAt places a click left of the text at the start of that line', async () => {
  const { caretTargetAt } = await loadCaret();
  const hit = caretTargetAt(boxes(), 40, 140);
  assert.equal(hit.index, 1);
  assert.equal(hit.x, 100.5);
});

test('caretTargetAt puts a click below the page at the end of the last paragraph', async () => {
  const { caretTargetAt } = await loadCaret();
  const hit = caretTargetAt(boxes(), 120, 900);
  assert.equal(hit.index, 2);
  assert.equal(hit.x, 699.5, 'the end of the last line, whatever the click column');
  assert.equal(hit.y, 171);
});

test('caretTargetAt uses the last line by position, not by document order', async () => {
  const { caretTargetAt } = await loadCaret();
  const reordered = [boxes()[2], boxes()[0]];
  const hit = caretTargetAt(reordered, 300, 900);
  assert.equal(hit.index, 0);
});

test('caretTargetAt reaches a click above the first line', async () => {
  const { caretTargetAt } = await loadCaret();
  const hit = caretTargetAt(boxes(), 150, 10);
  assert.equal(hit.index, 0);
  assert.equal(hit.y, 100.5);
});

test('caretTargetAt centres a point in a run too thin to inset', async () => {
  const { caretTargetAt } = await loadCaret();
  const hit = caretTargetAt([{ left: 200, right: 200, top: 100, bottom: 100 }], 400, 100);
  assert.deepEqual([hit.index, hit.x, hit.y], [0, 200, 100]);
});
