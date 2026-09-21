const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

async function loadEngineResult() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/engine-result.ts'),
    platform: 'node',
    transform: { jsx: { runtime: 'automatic' } },
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

test('a SEMANTIC_NO_OP refusal is success: no message, so no error banner', async () => {
  const { isSemanticNoOp, engineErrorMessage } = await loadEngineResult();
  const refusal = { ok: false, issues: [{ code: 'SEMANTIC_NO_OP', path: '/steps', message: 'transaction produces no native text change' }] };
  assert.equal(isSemanticNoOp(refusal), true);
  assert.equal(engineErrorMessage(refusal), '');
  assert.equal(engineErrorMessage({ code: 'SEMANTIC_NO_OP' }), '');
  assert.equal(engineErrorMessage(new Error('docx transaction rejected: SEMANTIC_NO_OP at /steps')), '');
  const wrapped = new Error('apply failed');
  wrapped.cause = { issues: [{ code: 'SEMANTIC_NO_OP' }] };
  assert.equal(engineErrorMessage(wrapped), '');
});

test('every other refusal keeps its message', async () => {
  const { isSemanticNoOp, engineErrorMessage } = await loadEngineResult();
  const refusal = { ok: false, issues: [{ code: 'STALE_REVISION', message: 'document moved on' }] };
  assert.equal(isSemanticNoOp(refusal), false);
  assert.equal(engineErrorMessage(new Error('document moved on')), 'document moved on');
  assert.equal(engineErrorMessage('plain failure'), 'plain failure');
  // A no-op alongside a real refusal is still a failure the user must see.
  const mixed = { issues: [{ code: 'SEMANTIC_NO_OP' }, { code: 'UNSUPPORTED_STRUCTURE' }] };
  assert.equal(isSemanticNoOp(mixed), false);
  assert.equal(engineErrorMessage(new Error('SEMANTIC_NO_OP_LOOKALIKE refused')), 'SEMANTIC_NO_OP_LOOKALIKE refused');
});
