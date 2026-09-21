const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

async function loadFormatting() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/formatting.ts'),
    platform: 'node',
    external: id => id.includes('packages/docs/src/') || id.includes('playground/src/') || id === '@injoffice/sheets/browser' || id === '@injoffice/pptx-native' || id === '@injoffice/pptx-wasm',
  });
  try {
    const { output } = await bundle.generate({ format: 'cjs', codeSplitting: false });
    const mod = { exports: {} };
    const req = id => {
      if (id.includes('docxRoundTrip')) return { editableDocxRuns: document => document._targets ?? [] };
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

function document() {
  const paragraph = {
    id: 'p1',
    edit_policy: { allowed_operations: ['properties.patch'] },
    properties: {},
    runs: [{ id: 'r1', kind: 'text', text: 'Hi', properties: { bold: true, font_family: 'Calibri', font_size_half_points: 22 }, anchor: { path: 'w:p/w:r[0]', part_name: 'word/document.xml' } }],
  };
  return {
    headers: [], footers: [], notes: [], comment_stories: [], sections: [],
    default_paragraph_style_id: 'Normal', default_run_properties: {}, default_paragraph_properties: {},
    paragraph_styles: [], numbering_definitions: [],
    _targets: [{ key: 'run-1', paragraphId: 'p1', runId: 'r1' }],
    body: { blocks: [{ paragraph }] },
  };
}

test('docxSelection and formattingValues expose run appearance for the toolbar', async () => {
  const { docxSelection, formattingValues } = await loadFormatting();
  const model = document();
  const selected = docxSelection(model, 'run-1');
  assert.equal(selected.paragraph.id, 'p1');
  assert.equal(selected.run.id, 'r1');
  const values = formattingValues({ kind: 'docx', document: model }, 'run-1');
  assert.equal(values.bold, true);
  assert.equal(values.font, 'Calibri');
  assert.equal(values.size, 11);
  assert.equal(formattingValues({ kind: 'docx', document: model }, 'missing'), undefined);
});


test('linked DOCX text retains toolbar values while disabling unsupported character formatting', async () => {
  const { formattingValues } = await loadFormatting();
  const model = document();
  model.body.blocks[0].paragraph.runs[0].hyperlink = {url: 'https://example.com', anchor: {}};
  const values = formattingValues({kind: 'docx', document: model}, 'run-1');
  assert.equal(values.bold, true);
  assert.equal(values.characterEditable, false);
});

test('DOCX highlight goes into the engine mutation and reads back into formatting values', async () => {
  const { documentFormatting, formattingValues } = await loadFormatting();
  const model = document(); model.source = { package_sha256: 'revision' };
  const request = documentFormatting(model, 'run-1', { highlight: 'yellow' }, 'highlight');
  assert.deepEqual(request.payload.mutations[0].properties, { highlight: 'yellow' });
  model.body.blocks[0].paragraph.runs[0].properties.highlight = 'yellow';
  assert.equal(formattingValues({ kind: 'docx', document: model }, 'run-1').highlight, 'yellow');
  assert.deepEqual(documentFormatting(model, 'run-1', { highlight: 'none' }, 'clear').payload.mutations[0].properties, { highlight: 'none' });
});
