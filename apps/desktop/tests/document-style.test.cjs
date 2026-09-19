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
