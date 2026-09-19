const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const React = require('react');
const { create, act } = require('react-test-renderer');
global.IS_REACT_ACT_ENVIRONMENT = true;

async function loadPreview() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/DocumentPreview.tsx'),
    platform: 'node',
    external: id => /^react(?:\/|$)/.test(id) || id.includes('packages/docs/src/') || id.includes('playground/src/'),
    transform: { jsx: { runtime: 'automatic' } },
  });
  try {
    const { output } = await bundle.generate({ format: 'cjs', codeSplitting: false });
    const mod = { exports: {} };
    const req = id => {
      if (id.includes('docxRoundTrip')) return { editableDocxRuns: document => document._targets ?? [] };
      if (id.includes('docsNativePreview')) return { nativeDocxRunText: run => run.text ?? '', nativeDocxParagraphText: () => '' };
      return require(id);
    };
    new Function('require', 'module', 'exports', output[0].code)(req, mod, mod.exports);
    return mod.exports.default ?? mod.exports;
  } finally {
    await bundle.close();
  }
}

function document() {
  return {
    headers: [], footers: [], notes: [], comment_stories: [],
    default_paragraph_style_id: 'Normal', default_run_properties: {}, default_paragraph_properties: {},
    paragraph_styles: [], numbering_definitions: [],
    _targets: [{ partName: 'word/document.xml', runId: 'r1', key: 'run-1' }],
    body: { blocks: [{ paragraph: {
      id: 'p1', can_format_range: true, anchor: { part_name: 'word/document.xml' }, properties: {},
      runs: [{ id: 'r1', kind: 'text', text: 'Hello', anchor: { part_name: 'word/document.xml' }, properties: {} }],
    } }] },
  };
}

test('DocumentPreview paints editable text, reports IME composition, and Escape cancels', async () => {
  const DocumentPreview = await loadPreview();
  global.window = { ...(global.window || {}), getSelection: () => null, document: { createTextNode: () => ({}) } };
  const composing = [];
  const cancelled = [];
  let view;
  await act(async () => {
    view = create(React.createElement(DocumentPreview, {
      images: {}, imageNotice: '', document: document(), selected: 'run-1', draft: 'Hello',
      onTextRangeChange() {}, joinPrevious() {}, insertLines() {}, busy: false, hasDraft: true,
      zoom: 100, navigation: false, choose() {}, updateDraft() {}, apply() {},
      cancel: () => cancelled.push('cancel'), onCompositionChange: value => composing.push(value), deleteImage() {},
    }));
  });
  const textbox = view.root.findByProps({ role: 'textbox' });
  assert.equal(textbox.props.contentEditable, 'plaintext-only');
  await act(async () => textbox.props.onCompositionStart());
  await act(async () => textbox.props.onCompositionEnd({ currentTarget: { textContent: 'Hello' } }));
  assert.deepEqual(composing, [true, false]);
  await act(async () => textbox.props.onKeyDown({ key: 'Escape', preventDefault() {}, nativeEvent: {} }));
  assert.deepEqual(cancelled, ['cancel']);
});
