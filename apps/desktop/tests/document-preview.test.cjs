const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const React = require('react');
const { create, act } = require('react-test-renderer');
global.IS_REACT_ACT_ENVIRONMENT = true;

async function loadPreview(exportName = 'default') {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/DocumentPreview.tsx'),
    platform: 'node',
    external: id => /^react(?:\/|$)/.test(id) || (id.includes('packages/docs/src/') && !id.includes('nativeNoteNumberingV1')) || id.includes('playground/src/'),
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
    return mod.exports[exportName] ?? mod.exports;
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

test('Enter and paragraph input split at the caret; Shift+Enter refuses a line break', async () => {
  const DocumentPreview = await loadPreview();
  const node = { textContent: 'Hello', contains: () => true, focus() {} };
  const range = {
    cloneRange() {
      let prefix = true;
      return { selectNodeContents() {}, setEnd() {}, setStart() { prefix = false; }, toString: () => prefix ? 'He' : 'llo' };
    },
    selectNodeContents() {}, collapse() {},
  };
  global.window = {
    getSelection: () => ({ rangeCount: 1, getRangeAt: () => range, removeAllRanges() {}, addRange() {} }),
    document: { createRange: () => range, addEventListener() {}, removeEventListener() {} },
  };
  const splits = [];
  let view;
  await act(async () => {
    view = create(React.createElement(DocumentPreview, {
      images: {}, imageNotice: '', document: document(), selected: 'run-1', draft: 'Hello',
      onTextRangeChange() {}, joinPrevious() {}, insertLines: (...args) => splits.push(args), busy: false, hasDraft: false,
      zoom: 100, navigation: false, choose() {}, updateDraft() {}, commit() {}, notice: '', onCompositionChange() {}, deleteImage() {},
    }), { createNodeMock: element => element.props.role === 'textbox' ? node : null });
  });
  try {
    const textbox = view.root.findByProps({ role: 'textbox' });
    const event = { key: 'Enter', preventDefault() {}, nativeEvent: {} };
    await act(async () => textbox.props.onKeyDown(event));
    await act(async () => textbox.props.onBeforeInput({ preventDefault() {}, nativeEvent: { inputType: 'insertParagraph' } }));
    assert.deepEqual(splits, [['He\nllo', 0], ['He\nllo', 0]]);
    await act(async () => textbox.props.onKeyDown({ ...event, shiftKey: true }));
    await act(async () => textbox.props.onBeforeInput({ preventDefault() {}, nativeEvent: { inputType: 'insertLineBreak' } }));
    assert.equal(splits.length, 2, 'line breaks must never become paragraph splits');
    assert.match(view.root.findByProps({ className: 'office-inline-hint' }).children[0], /Line breaks.*not supported.*DOCX engine/);
    await act(async () => textbox.props.onCompositionStart());
    await act(async () => textbox.props.onKeyDown(event));
    assert.equal(splits.length, 2, 'IME confirmation does not split');
  } finally { await act(async () => view.unmount()); }
});

test('flowing page metrics count trailing, repeated and overflow-adjacent hard breaks', async () => {
  const metrics = await loadPreview('flowingPageMetrics');
  assert.deepEqual(metrics(1000, 1000, 0, []), { page: 1, pages: 1 });
  assert.deepEqual(metrics(1000, 1000, 0, [200]), { page: 1, pages: 2 });
  assert.deepEqual(metrics(1000, 1000, 300, [200]), { page: 2, pages: 2 });
  assert.deepEqual(metrics(1000, 1000, 0, [200, 220]), { page: 1, pages: 3 });
  assert.deepEqual(metrics(2200, 1000, 2100, [2000]), { page: 3, pages: 3 });
  assert.deepEqual(metrics(2000, 1000, 0, [1000]), { page: 1, pages: 2 }, 'a break at a natural boundary is not double counted');
  assert.deepEqual(metrics(1000, 1000, 1000, [1000]), { page: 2, pages: 2 }, 'a final break creates an empty final page');
});


function noteStory(kind, id, role = 'content') {
  const part = `word/${kind}s.xml`;
  return { id: `${kind}-${id}`, native_story_id: String(id), kind, note_role: role, part_name: part,
    blocks: [{ paragraph: { id: `${kind}-p-${id}`, anchor: { part_name: part }, properties: {}, runs: [
      { id: `${kind}-r-${id}`, kind: 'text', text: `${kind} text ${id}`, anchor: { part_name: part }, properties: {} },
    ] } }] };
}
function noteReference(story, role = 'anchor') {
  return { id: `ref-${story.id}-${role}`, kind: 'reference', anchor: { part_name: 'word/document.xml' },
    reference: { kind: story.kind, target_id: story.id, role }, properties: {} };
}
async function renderNotes(value) {
  const DocumentPreview = await loadPreview();
  let view;
  await act(async () => { view = create(React.createElement(DocumentPreview, {
    images: {}, imageNotice: '', document: value, selected: '', draft: '',
    onTextRangeChange() {}, joinPrevious() {}, insertLines() {}, busy: false, hasDraft: false,
    zoom: 1, navigation: false, choose() {}, updateDraft() {}, commit() {}, notice: '', onCompositionChange() {}, deleteImage() {},
  })); });
  return view;
}

test('note sentinels and unreferenced content never appear as document notes', async () => {
  const value = document();
  value.notes = ['footnote', 'endnote'].flatMap(kind => [noteStory(kind, 0, 'separator'), noteStory(kind, 1, 'continuation-separator')]);
  for (const note of value.notes) note.blocks[0].paragraph.runs = [];
  let view = await renderNotes(value);
  assert.equal(view.root.findAllByProps({ className: 'office-story' }).length, 0);
  await act(async () => view.unmount());
  // Even malformed references to sentinels and label leaves are not content anchors.
  value.body.blocks[0].paragraph.runs.push(...value.notes.map(note => noteReference(note)));
  const orphan = noteStory('footnote', 99);
  value.notes.push(orphan);
  value.body.blocks[0].paragraph.runs.push(noteReference(orphan, 'label'));
  view = await renderNotes(value);
  assert.equal(view.root.findAllByProps({ className: 'office-story' }).length, 0);
  await act(async () => view.unmount());
});

test('two footnotes and an endnote follow body/table reference order with separate authored numbering', async () => {
  for (const authored of [false, true]) {
    const value = document();
    const first = noteStory('footnote', 42), second = noteStory('footnote', 7), end = noteStory('endnote', 42);
    value.notes = [second, noteStory('footnote', 0, 'separator'), end, noteStory('endnote', 1, 'continuation-separator'), first, noteStory('endnote', 99)];
    value.body.blocks[0].paragraph.runs.push(noteReference(first), noteReference(end));
    value.body.blocks.push({ table: { rows: [{ cells: [{ paragraphs: [{ id: 'table-p', anchor: { part_name: 'word/document.xml' }, properties: {}, runs: [noteReference(second)] }] }] }] } });
    for (const note of [first, second, end]) note.blocks[0].paragraph.runs.unshift(noteReference(note, 'label'));
    if (authored) value.note_numbering = [{ kind: 'footnote', format: 'upperLetter', labels: ['A', 'B'] }, { kind: 'endnote', format: 'lowerRoman', labels: ['i'] }];
    const view = await renderNotes(value);
    const markers = authored ? ['A', 'i', 'B'] : ['1', '1', '2'];
    assert.deepEqual(view.root.findAllByProps({ className: 'office-story' }).map(node => node.props['aria-label']), [`Footnote ${markers[0]}`, `Endnote ${markers[1]}`, `Footnote ${markers[2]}`]);
    assert.deepEqual(view.root.findAllByType('sup').map(node => node.children[0]), [...markers, ...markers]);
    await act(async () => view.unmount());
  }
});
