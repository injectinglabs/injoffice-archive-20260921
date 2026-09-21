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
      if (id.includes('docxPreviewImages')) return {extractDocxPreviewImages: async () => new Map()};
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
    headers: [], footers: [], notes: [], comment_stories: [], sections: [],
    default_paragraph_style_id: 'Normal', default_run_properties: {}, default_paragraph_properties: {},
    paragraph_styles: [], numbering_definitions: [],
    _targets: [{ partName: 'word/document.xml', runId: 'r1', paragraphId: 'p1', key: 'run-1' }],
    body: { blocks: [{ paragraph: {
      id: 'p1', can_format_range: true, anchor: { part_name: 'word/document.xml' }, properties: {},
      runs: [{ id: 'r1', kind: 'text', text: 'Hello', anchor: { part_name: 'word/document.xml' }, properties: { highlight: 'yellow' } }],
    } }] },
  };
}

test('DocumentPreview edits in place and commits without an apply step', async () => {
  const DocumentPreview = await loadPreview();
  global.window = {
    getSelection: () => null,
    document: {
      createTextNode: () => ({}),
      createRange: () => ({}),
      caretRangeFromPoint: () => null,
      addEventListener() {},
      removeEventListener() {},
    },
  };
  const composing = [];
  const committed = [];
  let view;
  await act(async () => {
    view = create(React.createElement(DocumentPreview, {
      images: {}, imageNotice: '', document: document(), selected: 'run-1', draft: 'Hello',
      onTextRangeChange() {}, joinPrevious() {}, insertLines() {}, busy: false, hasDraft: true,
      zoom: 100, navigation: false, choose() {}, updateDraft() {}, notice: '',
      commit: () => committed.push('commit'), onCompositionChange: value => composing.push(value), deleteImage() {},
    }));
  });
  const textbox = view.root.findByProps({ role: 'textbox' });
  assert.equal(textbox.props.contentEditable, 'plaintext-only');
  assert.equal(textbox.props.style.backgroundColor, '#FFFF00', 'engine highlighting is visible while editing');
  assert.equal(view.root.findAllByProps({ className: 'office-document-note' }).length, 0, 'no caption between the ribbon and the page');
  assert.equal(textbox.props['data-placeholder'], undefined, 'the edited paragraph carries no placeholder');
  await act(async () => textbox.props.onCompositionStart());
  await act(async () => textbox.props.onCompositionEnd({ currentTarget: { textContent: 'Hello' } }));
  assert.deepEqual(composing, [true, false]);
  // Esc leaves the paragraph by committing; there is no cancel.
  await act(async () => textbox.props.onKeyDown({ key: 'Escape', preventDefault() {}, nativeEvent: {} }));
  assert.deepEqual(committed, ['commit']);
  // Focus moving to another control belongs to that control; a click on nothing commits here.
  await act(async () => textbox.props.onBlur({ relatedTarget: {} }));
  assert.deepEqual(committed, ['commit']);
  await act(async () => textbox.props.onBlur({ relatedTarget: null }));
  assert.deepEqual(committed, ['commit', 'commit']);
  // Switching away from the window is not leaving the paragraph.
  global.window.document.hasFocus = () => false;
  await act(async () => textbox.props.onBlur({ relatedTarget: null }));
  assert.deepEqual(committed, ['commit', 'commit']);
  delete global.window.document.hasFocus;
});

test('DocumentPreview shows an engine notice beside the caret instead of a banner', async () => {
  const DocumentPreview = await loadPreview();
  global.window = {
    getSelection: () => null,
    document: {
      createTextNode: () => ({}),
      createRange: () => ({}),
      caretRangeFromPoint: () => null,
      addEventListener() {},
      removeEventListener() {},
    },
  };
  let view;
  const props = {
    images: {}, imageNotice: '', document: document(), selected: 'run-1', draft: 'Hello',
    onTextRangeChange() {}, joinPrevious() {}, insertLines() {}, busy: false, hasDraft: true,
    zoom: 100, navigation: false, choose() {}, updateDraft() {}, commit() {},
    onCompositionChange() {}, deleteImage() {}, notice: '',
  };
  await act(async () => { view = create(React.createElement(DocumentPreview, props)); });
  assert.equal(view.root.findAllByProps({ className: 'office-inline-hint' }).length, 0);
  await act(async () => { view.update(React.createElement(DocumentPreview, { ...props, notice: 'This paragraph cannot be split here yet.' })); });
  const hint = view.root.findByProps({ className: 'office-inline-hint' });
  assert.equal(hint.props.role, 'status');
  assert.equal(hint.children[0], 'This paragraph cannot be split here yet.');
});

function emptyDocument() {
  const value = document();
  value.body.blocks[0].paragraph.runs[0].text = '';
  return value;
}

function paperProps(view) {
  return view.root.findByProps({ className: 'office-paper' }).props;
}

function clickEvent(overrides) {
  const node = {
    dataset: { docxRun: 'r1' },
    getBoundingClientRect: () => ({ left: 100, right: 180, top: 100, bottom: 122 }),
    closest: () => ({ dataset: { docxParagraph: 'p1' } }),
    focus() { node.focused = true; },
  };
  const container = { querySelectorAll: () => [node] };
  const event = { button: 0, clientX: 600, clientY: 400, currentTarget: container, target: container, prevented: 0, preventDefault() { event.prevented++; }, ...overrides };
  event.node = node;
  return event;
}

test('DocumentPreview gives an empty paragraph the whole line as a click target', async () => {
  const DocumentPreview = await loadPreview();
  global.window = { getSelection: () => null, document: { createTextNode: () => ({}), createRange: () => ({}), caretRangeFromPoint: () => null, addEventListener() {}, removeEventListener() {} } };
  const props = {
    images: {}, imageNotice: '', document: emptyDocument(), selected: '', draft: '',
    onTextRangeChange() {}, joinPrevious() {}, insertLines() {}, busy: false, hasDraft: false,
    zoom: 100, navigation: false, choose() {}, updateDraft() {}, commit() {}, onCompositionChange() {}, deleteImage() {}, notice: '',
  };
  let view;
  await act(async () => { view = create(React.createElement(DocumentPreview, props)); });
  assert.equal(view.root.findByProps({ 'data-docx-run': 'r1' }).props.className, 'office-text-run office-text-run-fill');
  // A paragraph that has text keeps an ordinary inline run.
  await act(async () => { view.update(React.createElement(DocumentPreview, { ...props, document: document() })); });
  assert.equal(view.root.findByProps({ 'data-docx-run': 'r1' }).props.className, 'office-text-run');
  // The active run of an empty paragraph fills the line too, instead of a 2px span.
  await act(async () => { view.update(React.createElement(DocumentPreview, { ...props, document: emptyDocument(), selected: 'run-1' })); });
  assert.equal(view.root.findByProps({ role: 'textbox' }).props.className, 'office-inline-input office-inline-input-fill');
});

test('DocumentPreview places the caret from a click that missed the text', async () => {
  const DocumentPreview = await loadPreview();
  const added = [];
  global.window = {
    getSelection: () => ({ removeAllRanges() {}, addRange: range => added.push(range) }),
    document: { createTextNode: () => ({}), createRange: () => ({ selectNodeContents(node) { this.node = node; }, collapse(start) { this.collapsed = start; } }), caretRangeFromPoint: () => null, addEventListener() {}, removeEventListener() {} },
  };
  const chosen = [];
  const props = {
    images: {}, imageNotice: '', document: emptyDocument(), selected: '', draft: '',
    onTextRangeChange() {}, joinPrevious() {}, insertLines() {}, busy: false, hasDraft: false,
    zoom: 100, navigation: false, choose: key => chosen.push(key), updateDraft() {}, commit() {}, onCompositionChange() {}, deleteImage() {}, notice: '',
  };
  let view;
  await act(async () => { view = create(React.createElement(DocumentPreview, props)); });
  const event = clickEvent();
  await act(async () => paperProps(view).onMouseDown(event));
  assert.deepEqual(chosen, ['run-1'], 'the click below the paragraph selects its run');
  assert.equal(event.prevented, 1, 'the browser does not also place its own stray selection');
  // A click on a run, a button or an image belongs to that control.
  const onRun = clickEvent({ target: { closest: () => ({}) } });
  await act(async () => paperProps(view).onMouseDown(onRun));
  assert.deepEqual(chosen, ['run-1']);
  assert.equal(onRun.prevented, 0);
  // While the engine is working, forward the click so the editor can queue it.
  await act(async () => { view.update(React.createElement(DocumentPreview, { ...props, busy: true })); });
  await act(async () => paperProps(view).onMouseDown(clickEvent()));
  assert.deepEqual(chosen, ['run-1', 'run-1']);
  // The run is already the selected one: the caret moves in place, without re-choosing it.
  await act(async () => { view.update(React.createElement(DocumentPreview, { ...props, selected: 'run-1' })); });
  const again = clickEvent();
  await act(async () => paperProps(view).onMouseDown(again));
  assert.deepEqual(chosen, ['run-1', 'run-1']);
  assert.equal(again.prevented, 1);
  assert.equal(added.length, 1, 'the caret is placed in the active run');
  assert.equal(again.node.focused, true);
});
