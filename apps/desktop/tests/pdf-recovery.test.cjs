const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

async function loadRecovery() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/pdf-recovery.ts'),
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

function draft(extra = {}) {
  return { version: 1, format: 'pdf', page: 1, tool: 'text', text: 'Hi', size: 12, color: '#112233', placement: { at: [10, 20] }, ...extra };
}

test('PDF recovery drafts accept a text note and reject unknown fields or invalid geometry', async () => {
  const { parsePdfRecoveryDraft } = await loadRecovery();
  const parsed = parsePdfRecoveryDraft(draft());
  assert.equal(parsed.tool, 'text');
  assert.equal(parsed.text, 'Hi');
  assert.deepEqual(parsed.placement.at, [10, 20]);
  assert.equal(parsePdfRecoveryDraft({ ...draft(), extra: true }), null);
  assert.equal(parsePdfRecoveryDraft({ ...draft(), color: 'red' }), null);
  assert.equal(parsePdfRecoveryDraft({ ...draft(), page: 0 }), null);
  assert.equal(parsePdfRecoveryDraft({ ...draft(), tool: 'highlight', text: 'no' }), null);
  assert.equal(parsePdfRecoveryDraft(Object.create({ version: 1 })), null);
});
