const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

async function loadStyle() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/document-style.ts'),
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

function documentWithStyles(styles, paragraphStyle, runProperties = {}) {
  return {
    default_paragraph_style_id: 'Normal',
    default_run_properties: { bold: false, italic: false },
    default_paragraph_properties: {},
    paragraph_styles: styles,
    numbering_definitions: [],
    body: { blocks: [] },
    _paragraph: { properties: { paragraph_style_id: paragraphStyle }, runs: [{ properties: runProperties }] },
  };
}

test('paragraphAppearance toggles OOXML style bold and refuses cycles', async () => {
  const { paragraphAppearance, runAppearance } = await loadStyle();
  const styles = [
    { id: 'Normal', paragraph_properties: {}, run_properties: {} },
    { id: 'Heading1', based_on: 'Normal', paragraph_properties: { spacing_after_twips: 240 }, run_properties: { bold: true } },
    { id: 'Emphasis', based_on: 'Heading1', run_properties: { bold: true, italic: true } },
    { id: 'Loop', based_on: 'Loop', run_properties: { bold: true } },
  ];
  const heading = documentWithStyles(styles, 'Heading1');
  const headingLook = paragraphAppearance(heading, heading._paragraph);
  assert.equal(headingLook.complete, true);
  assert.equal(headingLook.run.bold, true);
  assert.equal(headingLook.paragraph.spacing_after_twips, 240);
  assert.equal(runAppearance(heading, heading._paragraph, heading._paragraph.runs[0]).bold, true);

  const emphasis = documentWithStyles(styles, 'Emphasis');
  assert.equal(paragraphAppearance(emphasis, emphasis._paragraph).run.bold, false);
  assert.equal(paragraphAppearance(emphasis, emphasis._paragraph).run.italic, true);

  const loop = documentWithStyles(styles, 'Loop');
  assert.equal(paragraphAppearance(loop, loop._paragraph).complete, false);
  assert.equal(paragraphAppearance(loop, loop._paragraph).run.bold, false);
});

// canFormatRun mirrors what the native engine will accept, so the ribbon can
// grey a control out instead of letting the engine refuse the click.
test('canFormatRun and canFormatParagraphRange mirror the engine\'s own refusals', async () => {
  const { canFormatRun, canFormatParagraphRange } = await loadStyle();
  const policy = { mode: 'read-write', allowed_operations: ['text.replace', 'properties.patch'] };
  const anchor = (path, part = 'word/document.xml') => ({ part_name: part, path, start_byte: 0, end_byte: 1, xml_sha256: 'sha256:' + '0'.repeat(64) });
  const paragraphPath = '/w:document[1]/w:body[1]/w:p[1]';
  const run = (path, kind = 'text') => ({ kind, id: `run:${path}`, anchor: anchor(`${paragraphPath}${path}`), text: kind === 'text' ? 'text' : undefined });
  const paragraph = (runs, edit_policy = policy) => ({ id: 'paragraph:1', anchor: anchor(paragraphPath), edit_policy, properties: {}, runs });

  const plain = paragraph([run('/w:r[1]/w:t[1]')]);
  assert.equal(canFormatRun(plain, plain.runs[0]), true);
  assert.equal(canFormatParagraphRange(plain), true);

  const textOnly = paragraph([{ ...run('/w:r[1]/w:t[1]'), kind: 'control', control: 'tab', text: undefined }]);
  assert.equal(canFormatRun(textOnly, textOnly.runs[0]), false, 'a control run carries no text to format');
  assert.equal(canFormatParagraphRange(textOnly), false, 'a paragraph with no text run has nothing to format');

  const shared = paragraph([run('/w:r[1]/w:t[1]'), { ...run('/w:r[1]/w:tab[1]', 'control'), control: 'tab' }]);
  assert.equal(canFormatRun(shared, shared.runs[0]), false, 'the engine cannot split a run that also holds a tab');
  assert.equal(canFormatParagraphRange(shared), false);

  const hyperlinked = paragraph([run('/w:hyperlink[1]/w:r[1]/w:t[1]')]);
  assert.equal(canFormatRun(hyperlinked, hyperlinked.runs[0]), false, 'a run under w:hyperlink is not a direct child');
  assert.equal(canFormatParagraphRange(hyperlinked), false);

  const readOnly = paragraph([run('/w:r[1]/w:t[1]')], { mode: 'read-only', allowed_operations: [], refusal: { code: 'UNMODELED_PARAGRAPH_MARKUP', message: 'x', preservation: 'refuse-mutation' } });
  assert.equal(canFormatParagraphRange(readOnly), false);
  const textOnlyPolicy = paragraph([run('/w:r[1]/w:t[1]')], { mode: 'read-write', allowed_operations: ['text.replace'] });
  assert.equal(canFormatParagraphRange(textOnlyPolicy), false, 'a paragraph that only admits text.replace is not formattable');
  const { edit_policy: _dropped, ...policyless } = paragraph([run('/w:r[1]/w:t[1]')]);
  assert.equal(canFormatParagraphRange(policyless), false, 'an absent policy is a refusal');

  const twoRuns = paragraph([run('/w:r[1]/w:t[1]'), run('/w:r[2]/w:t[1]')]);
  assert.equal(canFormatParagraphRange(twoRuns), true, 'a selection may span sibling runs');
});
