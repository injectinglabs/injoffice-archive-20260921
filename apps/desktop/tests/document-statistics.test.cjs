const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

async function loadStats() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/document-statistics.ts'),
    platform: 'node',
    external: id => id.includes('packages/docs/src/'),
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

function paragraph(id, runs) {
  return { id, runs: runs.map((text, index) => ({ id: `${id}-r${index}`, kind: 'text', text })) };
}

function document(paragraphs) {
  return { body: { blocks: paragraphs.map(item => ({ paragraph: item })) } };
}

test('documentStatistics counts words and a draft keystroke does not recount other paragraphs', async () => {
  const { documentStatistics, statisticsWithDraft } = await loadStats();
  const first = paragraph('p1', ['Hello world']);
  const second = paragraph('p2', ['More']);
  const base = documentStatistics(document([first, second]));
  assert.equal(base.words, 3);
  assert.equal(base.characters, 'Hello world'.length + 'More'.length);
  assert.equal(base.paragraphs, 2);
  const draft = statisticsWithDraft(base, { paragraphId: 'p1', runId: 'p1-r0', text: 'Hello worlds' });
  assert.equal(draft.words, 3);
  assert.equal(draft.characters, base.characters + 1);
  assert.equal(draft.paragraphs, 2);
  const extra = statisticsWithDraft(base, { paragraphId: 'p1', runId: 'p1-r0', text: 'Hello world extra' });
  assert.equal(extra.words, 4);
});
