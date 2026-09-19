const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

async function loadMode() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/presentationMode.ts'),
    platform: 'node',
    external: id => id === '@injoffice/pptx-native',
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

test('presentation mode owns navigation, clamps the start index, and Escape exits', async () => {
  const { startPresentationMode, navigatePresentation, presentationScale } = await loadMode();
  const deck = { slides: [{}, {}, {}] };
  const started = startPresentationMode(deck, 99, false);
  assert.equal(started.index, 2);
  assert.equal(navigatePresentation(started, 'Home').index, 0);
  assert.equal(navigatePresentation(started, 'ArrowLeft').index, 1);
  assert.equal(navigatePresentation(started, 'Escape'), null);
  assert.throws(() => startPresentationMode({ slides: [] }, 0, false), /no slides/);
  assert.equal(presentationScale(1920, 1080, 1920, 1080), 1);
  assert.equal(presentationScale(960, 1080, 1920, 1080), 0.5);
  assert.equal(presentationScale(0, 1080, 1920, 1080), 0);
});
