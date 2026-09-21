const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

async function loadAuthoring() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/document-authoring.ts'),
    platform: 'node',
    external: id => id.includes('packages/docs/src/') || id.includes('playground/src/') || id === '@injoffice/sheets/browser' || id === '@injoffice/pptx-native' || id === '@injoffice/pptx-wasm' || id === '@injoffice/docx-wasm',
  });
  try {
    const { output } = await bundle.generate({ format: 'cjs', codeSplitting: false });
    const mod = { exports: {} };
    const req = id => {
      if (id.includes('docxPreviewImages')) return {extractDocxPreviewImages: async () => new Map()};
      if (id.includes('docxRoundTrip')) return {
        editableDocxRuns: document => document._targets ?? [],
        buildDocxRunMutation: (document, target, text, mutationId) => ({
          protocol: 'injoffice.office.mutations', version: 1, format: 'docx', mutation_id: mutationId,
          expected_revision: document.source.package_sha256,
          payload: { mutations: [{ target_kind: 'run', target_id: target.runId, operation: 'text.replace', text }] },
        }),
      };
      if (id.includes('nativeRoundTrip')) return { editableTargets: () => [], targetKey: () => '' };
      if (id.includes('pptxRoundTrip')) return { editablePptxTextTargets: () => [], pptxTargetKey: () => '' };
      return require(id);
    };
    new Function('require', 'module', 'exports', output[0].code)(req, mod, mod.exports);
    return mod.exports;
  } finally {
    await bundle.close();
  }
}

function paragraph(id, text, extra = {}) {
  return {
    id,
    edit_policy: { allowed_operations: ['block.insert_after', 'paragraph.split', 'paragraph.join_previous', 'properties.patch'] },
    properties: {},
    anchor: { xml_sha256: `sha-${id}` },
    runs: [{ id: `${id}-r0`, kind: 'text', text, properties: {}, anchor: { path: 'w:p/w:r[0]', part_name: 'word/document.xml' } }],
    ...extra,
  };
}

function document(paragraphs) {
  const blocks = paragraphs.map(item => ({ paragraph: item }));
  const targets = paragraphs.map(item => ({
    key: item.id,
    paragraphId: item.id,
    runId: item.runs[0].id,
    partName: 'word/document.xml',
    text: item.runs[0].text,
  }));
  return {
    source: { package_sha256: 'rev' },
    headers: [], footers: [], notes: [], comment_stories: [], sections: [],
    default_paragraph_style_id: 'Normal', default_run_properties: {}, default_paragraph_properties: {},
    paragraph_styles: [], numbering_definitions: [],
    _targets: targets,
    body: { blocks },
  };
}

function client(extract = model => model) {
  const applied = [];
  return {
    applied,
    apply: async (bytes, model, envelope) => {
      applied.push(envelope);
      return Buffer.from('next');
    },
    extract: async () => extract(),
  };
}

test('replaceEditableDocumentText refuses bad queries and applies a single-line replacement', async () => {
  const { replaceEditableDocumentText, insertDocumentTable, replaceParagraphLines } = await loadAuthoring();
  const model = document([paragraph('p1', 'Hello world')]);
  await assert.rejects(replaceEditableDocumentText({}, Buffer.from('x'), model, '', 'a', () => 'id'), /different single-line/);
  await assert.rejects(replaceEditableDocumentText({}, Buffer.from('x'), model, 'Hello', 'Hello', () => 'id'), /different single-line/);
  await assert.rejects(replaceEditableDocumentText({}, Buffer.from('x'), model, 'Hello\n', 'Hi', () => 'id'), /different single-line/);
  await assert.rejects(replaceEditableDocumentText({}, Buffer.from('x'), model, 'missing', 'Hi', () => 'id'), /No matching/);
  const next = document([paragraph('p1', 'Hey world')]);
  next._targets[0].text = 'Hey world';
  const wasm = client(() => next);
  const result = await replaceEditableDocumentText(wasm, Buffer.from('x'), model, 'Hello', 'Hey', () => 'mut-1');
  assert.equal(result.document._targets[0].text, 'Hey world');
  assert.equal(wasm.applied[0].payload.mutations[0].text, 'Hey world');

  await assert.rejects(insertDocumentTable({}, Buffer.from('x'), model, 'p1', 0, 2, () => 'id'), /1–20 rows/);
  await assert.rejects(insertDocumentTable({}, Buffer.from('x'), model, 'p1', 2, 13, () => 'id'), /1–12 columns/);
  await assert.rejects(replaceParagraphLines({}, Buffer.from('x'), model, 'p1', 'a\tb', () => 'id'), /without tabs/);
});

test('page-break insertion targets the caret and selects the re-extracted suffix', async () => {
  const { insertDocumentPageBreak, canInsertDocumentPageBreak } = await loadAuthoring();
  const sourceParagraph = paragraph('p1', 'BeforeAfter');
  sourceParagraph.edit_policy.allowed_operations.push('page_break.insert');
  sourceParagraph.runs[0].anchor.path = 'w:p/w:r[1]/w:t[1]';
  const model = document([sourceParagraph]);
  assert.equal(canInsertDocumentPageBreak(model, 'p1'), true);
  const nextParagraph = paragraph('p1', 'Before');
  nextParagraph.runs.push(
    { id: 'break', kind: 'control', control: 'page-break', anchor: { path: 'w:p/w:r[1]/w:br[1]', part_name: 'word/document.xml' } },
    { id: 'suffix', kind: 'text', text: 'After', anchor: { path: 'w:p/w:r[1]/w:t[2]', part_name: 'word/document.xml' } },
  );
  const next = document([nextParagraph]);
  next._targets.push({ key: 'suffix-key', paragraphId: 'p1', runId: 'suffix', partName: 'word/document.xml', text: 'After' });
  const wasm = client(() => next);
  const result = await insertDocumentPageBreak(wasm, Buffer.from('source'), model, 'p1', 6, () => 'break-1');
  assert.equal(result.key, 'suffix-key');
  assert.equal(result.text, 'After');
  assert.deepEqual(wasm.applied[0].payload.mutations[0].split, { run_id: 'p1-r0', offset_utf16: 6 });
  assert.equal(wasm.applied[0].payload.mutations[0].operation, 'page_break.insert');
  sourceParagraph.runs[0].hyperlink = { url: 'https://example.com' };
  assert.equal(canInsertDocumentPageBreak(model, 'p1'), false);
  await assert.rejects(insertDocumentPageBreak(wasm, Buffer.from('source'), model, 'p1', 6, () => 'id'), /supported body text/);
  assert.equal(wasm.applied.length, 1);
});
