const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const React = require('react');
const { create, act } = require('react-test-renderer');
global.IS_REACT_ACT_ENVIRONMENT = true;

/** The editor's own idle-commit delay, read from the module under test. */
let IDLE_COMMIT_MS;

async function loadEditor(client) {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/OfficeEditor.tsx'),
    platform: 'node',
    external: id => /^react(?:\/|$)/.test(id) || id === '@injoffice/docx-wasm' || id.includes('playground/src/') || id.includes('packages/docs/src/') || id === '@injoffice/sheets/browser' || id === '@injoffice/pptx-native' || id === '@injoffice/pptx-wasm',
    transform: { jsx: { runtime: 'automatic' } },
    plugins: [{
      name: 'css',
      resolveId(id) { if (id.endsWith('.css')) return '\0css'; },
      load(id) { if (id === '\0css') return 'export default ""'; },
    }],
  });
  try {
    const { output } = await bundle.generate({ format: 'cjs', codeSplitting: false });
    const mod = { exports: {} };
    const req = id => {
      if (id === '@injoffice/docx-wasm') return { createDocxWasmClient: () => client };
      if (id.includes('docxPreviewImages')) return { extractDocxPreviewImages: async () => new Map() };
      if (id.includes('docxRoundTrip')) return {
        editableDocxRuns: document => document._targets ?? [],
        buildDocxRunMutation: (document, target, text, mutationId) => ({
          protocol: 'injoffice.office.mutations', version: 1, format: 'docx', mutation_id: mutationId,
          expected_revision: document.source.package_sha256,
          payload: { mutations: [{ target_kind: 'run', target_id: target.runId, operation: 'text.replace', text }] },
        }),
        verifyDocxRoundTrip() {},
      };
      if (id.includes('docsNativePreview')) return { nativeDocxRunText: run => run.text ?? '', nativeDocxParagraphText: () => '' };
      if (id.includes('nativeRoundTrip')) return { editableTargets: () => [], targetKey: () => '' };
      if (id.includes('pptxRoundTrip')) return { editablePptxTextTargets: () => [], pptxTargetKey: () => '' };
      return require(id);
    };
    new Function('require', 'module', 'exports', output[0].code)(req, mod, mod.exports);
    IDLE_COMMIT_MS = mod.exports.IDLE_COMMIT_MS;
    assert.equal(typeof IDLE_COMMIT_MS, 'number');
    return mod.exports.default ?? mod.exports;
  } finally {
    await bundle.close();
  }
}

async function loadModule(file) {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({ input: path.resolve(__dirname, '../src', file), platform: 'node', external: id => /^react(?:\/|$)/.test(id) });
  try {
    const { output } = await bundle.generate({ format: 'cjs', codeSplitting: false });
    const mod = { exports: {} };
    new Function('require', 'module', 'exports', output[0].code)(require, mod, mod.exports);
    return mod.exports;
  } finally {
    await bundle.close();
  }
}

function tinyDocument(text = '') {
  const paragraph = {
    id: 'p1', can_format_range: true,
    edit_policy: { mode: 'read-write', allowed_operations: ['properties.patch', 'block.insert_after'] },
    anchor: { part_name: 'word/document.xml', path: 'w:p[0]', xml_sha256: 'p1-sha' },
    properties: {},
    runs: [{
      id: 'r1', kind: 'text', text, can_format_range: true,
      anchor: { part_name: 'word/document.xml', path: 'w:p[0]/w:r[0]', xml_sha256: 'r1-sha' },
      properties: {},
    }],
  };
  return {
    source: { package_sha256: 'rev' },
    headers: [], footers: [], notes: [], comment_stories: [], sections: [],
    default_paragraph_style_id: 'Normal', default_run_properties: {}, default_paragraph_properties: {},
    paragraph_styles: [], numbering_definitions: [],
    _targets: [{ key: 'run-1', label: 'Body', text, partName: 'word/document.xml', runId: 'r1', paragraphId: 'p1' }],
    body: { blocks: [{ paragraph }] },
  };
}

function mockClient(model) {
  const applied = [];
  return {
    applied,
    extract: async () => model,
    apply: async (bytes, _document, envelope) => {
      applied.push(envelope);
      const text = envelope?.payload?.mutations?.[0]?.text;
      if (typeof text === 'string') {
        model._targets[0].text = text;
        model.body.blocks[0].paragraph.runs[0].text = text;
      }
      return bytes;
    },
    terminate() { applied.push('terminate'); },
  };
}

async function until(predicate) {
  for (let count = 0; count < 200; count++) {
    if (predicate()) return;
    await act(async () => new Promise(resolve => setTimeout(resolve, 5)));
  }
  throw new Error('Editor did not settle');
}

function mockWindow(getSelection = () => null) {
  global.window = {
    getSelection,
    document: {
      createTextNode: () => ({}),
      createRange: () => ({}),
      caretRangeFromPoint: () => null,
      addEventListener() {},
      removeEventListener() {},
    },
  };
}

function runPrefixSelection(prefix) {
  const run = { closest(sel) { return sel === '[data-docx-run]' ? run : null; } };
  const textNode = { parentElement: run };
  return {
    rangeCount: 1,
    getRangeAt: () => ({
      startContainer: textNode,
      startOffset: prefix.length,
      cloneRange() {
        return {
          selectNodeContents() {},
          setEnd() {},
          toString() { return prefix; },
        };
      },
    }),
  };
}

function nativeApplies(client) {
  return client.applied.filter(item => item && item.payload);
}

async function mountEditor(client, props = {}) {
  const OfficeEditor = await loadEditor(client);
  const changes = [];
  const busy = [];
  let view;
  await act(async () => {
    view = create(React.createElement(OfficeEditor, {
      name: 'Note.docx',
      bytes: new Uint8Array([0x50, 0x4b]),
      onChange: value => changes.push(value),
      onBusyChange: value => busy.push(value),
      ...props,
    }));
  });
  await until(() => view.root.findAllByProps({ 'aria-label': 'Document content' }).length > 0);
  return { view, changes, busy };
}

async function settleApply(changes, busy) {
  for (let count = 0; count < 50 && (changes.length === 0 || busy.at(-1) !== false); count++) {
    await act(async () => Promise.resolve());
  }
}

test('OfficeEditor mounts with mocked wasm, reports busy, and applies a draft', async () => {
  assert.equal(fs.existsSync(path.resolve(__dirname, '../src/office-editor.css')), true);
  const bytes = new Uint8Array([0x50, 0x4b]);
  const client = mockClient(tinyDocument(''));
  const OfficeEditor = await loadEditor(client);
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
  const busy = [];
  const changes = [];
  let view;
  await act(async () => {
    view = create(React.createElement(OfficeEditor, {
      name: 'Note.docx',
      bytes,
      onChange: value => changes.push(value),
      onBusyChange: value => busy.push(value),
    }));
  });
  await until(() => view.root.findAllByProps({ 'aria-label': 'Document content' }).length > 0);
  assert.equal(view.root.findByProps({ className: 'office-editor office-editor-document' }) != null, true);
  assert.equal(busy.includes(true), true);
  assert.equal(busy.at(-1), false);
  const textbox = view.root.findByProps({ role: 'textbox' });
  assert.equal(view.root.findAllByProps({ className: 'office-apply' }).length, 0, 'Word has no apply step');
  await act(async () => textbox.props.onInput({ currentTarget: { textContent: 'Hello' } }));
  // Leaving the paragraph commits through the same engine path the Apply button used.
  await act(async () => textbox.props.onBlur({ relatedTarget: null }));
  await until(() => changes.length === 1 && busy.at(-1) === false);
  assert.equal(changes[0], bytes);
  assert.equal(client.applied[0].payload.mutations[0].text, 'Hello');
  await act(async () => view.unmount());
});

function documentPreview(view) {
  return view.root.find(node => typeof node.props?.choose === 'function' && typeof node.props?.updateDraft === 'function');
}

test('OfficeEditor schedules a hidden native apply after debounce and skips while composing', async t => {
  const bytes = new Uint8Array([0x50, 0x4b]);
  const client = mockClient(tinyDocument(''));
  const OfficeEditor = await loadEditor(client);
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
  const changes = [];
  const busy = [];
  let view;
  await act(async () => {
    view = create(React.createElement(OfficeEditor, {
      name: 'Note.docx',
      bytes,
      onChange: value => changes.push(value),
      onBusyChange: value => busy.push(value),
    }));
  });
  await until(() => view.root.findAllByProps({ 'aria-label': 'Document content' }).length > 0);
  t.mock.timers.enable({ apis: ['setTimeout'] });
  const textbox = view.root.findByProps({ role: 'textbox' });
  await act(async () => textbox.props.onInput({ currentTarget: { textContent: 'Hello' } }));
  assert.equal(client.applied.length, 0);
  assert.equal(changes.length, 0);
  await act(async () => { t.mock.timers.tick(IDLE_COMMIT_MS - 1); });
  assert.equal(client.applied.length, 0);
  await act(async () => { t.mock.timers.tick(1); });
  for (let count = 0; count < 50 && (changes.length === 0 || busy.at(-1) !== false); count++) {
    await act(async () => Promise.resolve());
  }
  assert.equal(client.applied[0].payload.mutations[0].text, 'Hello');
  assert.equal(changes.length, 1);
  assert.equal(view.root.findAllByProps({ role: 'textbox' }).length, 1);
  assert.equal(documentPreview(view).props.caretOffset, 5);
  const live = view.root.findByProps({ role: 'textbox' });
  await act(async () => live.props.onCompositionStart());
  await act(async () => live.props.onInput({ currentTarget: { textContent: 'Hello!' } }));
  await act(async () => { t.mock.timers.tick(IDLE_COMMIT_MS); });
  for (let count = 0; count < 10; count++) await act(async () => Promise.resolve());
  assert.equal(client.applied.length, 1);
  await act(async () => view.unmount());
});

test('OfficeEditor compositionend schedules hidden apply after debounce', async t => {
  const client = mockClient(tinyDocument(''));
  mockWindow();
  const { view, changes, busy } = await mountEditor(client);
  t.mock.timers.enable({ apis: ['setTimeout'] });
  const textbox = view.root.findByProps({ role: 'textbox' });
  await act(async () => textbox.props.onCompositionStart());
  await act(async () => textbox.props.onInput({ currentTarget: { textContent: 'Hello' } }));
  await act(async () => { t.mock.timers.tick(IDLE_COMMIT_MS); });
  for (let count = 0; count < 10; count++) await act(async () => Promise.resolve());
  assert.equal(nativeApplies(client).length, 0);
  await act(async () => textbox.props.onCompositionEnd({ currentTarget: { textContent: 'Hello' } }));
  assert.equal(nativeApplies(client).length, 0);
  await act(async () => { t.mock.timers.tick(IDLE_COMMIT_MS - 1); });
  assert.equal(nativeApplies(client).length, 0);
  await act(async () => { t.mock.timers.tick(1); });
  await settleApply(changes, busy);
  assert.equal(nativeApplies(client)[0].payload.mutations[0].text, 'Hello');
  await act(async () => view.unmount());
});

test('OfficeEditor hidden apply restores a non-end caret from the run prefix range', async t => {
  const client = mockClient(tinyDocument(''));
  mockWindow(() => runPrefixSelection('He'));
  const { view, changes, busy } = await mountEditor(client);
  t.mock.timers.enable({ apis: ['setTimeout'] });
  const textbox = view.root.findByProps({ role: 'textbox' });
  await act(async () => textbox.props.onInput({ currentTarget: { textContent: 'Hello' } }));
  await act(async () => { t.mock.timers.tick(IDLE_COMMIT_MS); });
  await settleApply(changes, busy);
  assert.equal(nativeApplies(client)[0].payload.mutations[0].text, 'Hello');
  assert.equal(documentPreview(view).props.caretOffset, 2);
  await act(async () => view.unmount());
});

test('OfficeEditor commit applies a pending debounce once', async t => {
  const client = mockClient(tinyDocument(''));
  mockWindow();
  let commit = async () => false;
  const { view, changes, busy } = await mountEditor(client, {
    registerCommit: fn => { commit = fn; },
  });
  t.mock.timers.enable({ apis: ['setTimeout'] });
  const textbox = view.root.findByProps({ role: 'textbox' });
  await act(async () => textbox.props.onInput({ currentTarget: { textContent: 'Hello' } }));
  assert.equal(nativeApplies(client).length, 0);
  await act(async () => { await commit(); });
  await settleApply(changes, busy);
  assert.equal(nativeApplies(client).length, 1);
  assert.equal(nativeApplies(client)[0].payload.mutations[0].text, 'Hello');
  await act(async () => { t.mock.timers.tick(IDLE_COMMIT_MS); });
  for (let count = 0; count < 10; count++) await act(async () => Promise.resolve());
  assert.equal(nativeApplies(client).length, 1);
  await act(async () => view.unmount());
});

test('OfficeEditor answers a break the paragraph cannot take with a caret notice, not a banner', async t => {
  const client = mockClient(tinyDocument(''));
  mockWindow();
  const { view, changes, busy } = await mountEditor(client);
  try {
    t.mock.timers.enable({ apis: ['setTimeout'] });
    await act(async () => view.root.findByProps({ role: 'textbox' }).props.onInput({ currentTarget: { textContent: 'Hello' } }));
    // This paragraph has no paragraph.split policy: the engine refuses the break.
    await act(async () => { await documentPreview(view).props.insertLines('Hel\nlo', 2); });
    for (let count = 0; count < 20 && documentPreview(view).props.notice === ''; count++) await act(async () => Promise.resolve());
    assert.match(documentPreview(view).props.notice, /paragraph/i);
    assert.equal(view.root.findAllByProps({ className: 'office-error' }).length, 0, 'a refused break is not an error banner');
    // The typed text is not lost: the idle commit still writes it.
    await act(async () => { t.mock.timers.tick(IDLE_COMMIT_MS); });
    await settleApply(changes, busy);
    assert.equal(nativeApplies(client).at(-1).payload.mutations[0].text, 'Hello');
  } finally {
    await act(async () => view.unmount());
  }
});

test('OfficeEditor carries one Word-like status row: page and words left, style right', async () => {
  const client = mockClient(tinyDocument('Hello there'));
  mockWindow();
  const { view } = await mountEditor(client);
  const text = node => node.children.map(child => typeof child === 'string' ? child : text(child)).join('');
  try {
    const status = view.root.findByProps({ 'aria-label': 'Document status' });
    const spans = status.findAllByType('span').map(text);
    assert.equal(spans[0], 'Page 1 of 1 · 2 words');
    assert.equal(spans.at(-1), 'Normal', 'the paragraph style sits on the right, as in Word');
    const info = status.findByProps({ className: 'office-status-info' });
    assert.match(info.props.title, /11 characters/);
    assert.match(info.props.title, /original page layout is preserved in the file/);
    assert.equal(/character|Saved|Unsaved/.test(spans.join(' ')), false, 'no character count or duplicated save state in the row');
    assert.equal(view.root.findAllByProps({ className: 'office-document-note' }).length, 0, 'no caption above the page');
  } finally {
    await act(async () => view.unmount());
  }
});

test('OfficeEditor says in the status row when no paragraph can be edited', async () => {
  const model = tinyDocument('Fixed text');
  model._targets = [];
  const client = mockClient(model);
  mockWindow();
  const { view } = await mountEditor(client);
  const text = node => node.children.map(child => typeof child === 'string' ? child : text(child)).join('');
  try {
    const alert = view.root.findByProps({ className: 'office-status-alert' });
    assert.equal(alert.props.role, 'status');
    assert.match(text(alert), /no editable paragraphs/i);
    assert.match(alert.props.title, /caret/i, 'the row explains why clicking the page does nothing');
  } finally {
    await act(async () => view.unmount());
  }
});

test('OfficeEditor keeps the status row clear while a paragraph is editable', async () => {
  const client = mockClient(tinyDocument('Hello there'));
  mockWindow();
  const { view } = await mountEditor(client);
  try {
    assert.equal(view.root.findAllByProps({ className: 'office-status-alert' }).length, 0);
  } finally {
    await act(async () => view.unmount());
  }
});

test('OfficeEditor arranges its controls as a Word ribbon with labelled groups, icons, and shortcut tooltips', async () => {
  const { shortcutLabel } = await loadModule('shortcuts.ts');
  const client = mockClient(tinyDocument('Hello'));
  const { view } = await mountEditor(client);
  const text = node => node.children.map(child => typeof child === 'string' ? child : text(child)).join('');
  try {
    const tabs = view.root.findAllByProps({ role: 'tab' }).map(text);
    assert.deepEqual(tabs, ['Home', 'Insert', 'Layout'], 'File is omitted without a PDF export bridge; no tab is ever empty');
    const panels = view.root.findAllByProps({ role: 'tabpanel' });
    const groups = panel => panel.findAllByProps({ role: 'group' }).map(group => group.props['aria-label']);
    assert.deepEqual(groups(panels[0]), ['Font', 'Paragraph', 'Styles', 'Editing']);
    assert.deepEqual(groups(panels[1]), ['Pages', 'Tables', 'Links', 'Text']);
    assert.deepEqual(groups(panels[2]), ['Paragraph'], 'Page Setup is dropped when the document has no single section');
    assert.equal(view.root.findAllByProps({ role: 'toolbar' }).filter(toolbar => toolbar.props['aria-label'] === 'Quick access').length, 0, 'Undo/Redo live in the shell title bar, not in the ribbon');
    const commandButtons = view.root.findAllByProps({ role: 'group' }).flatMap(group => group.findAllByType('button'));
    assert.ok(commandButtons.length >= 8, `command buttons: ${commandButtons.length}`);
    for (const button of commandButtons) {
      assert.equal(button.findAllByType('svg').length, 1, `${button.props.title} has an icon`);
      assert.ok(button.props.title, 'every command button has a tooltip');
    }
    const titles = Object.fromEntries(commandButtons.map(button => [button.props['aria-label'] ?? text(button), button.props.title]));
    assert.equal(titles.Bold, `Select editable text to format. (${shortcutLabel('bold')})`);
    assert.equal(titles.Italic, `Select editable text to format. (${shortcutLabel('italic')})`);
    assert.equal(titles.Underline, `Select editable text to format. (${shortcutLabel('underline')})`);
    assert.equal(titles['Find / replace'], `Find / replace (${shortcutLabel('find')})`);
    assert.equal(view.root.findAllByProps({ className: 'office-toolbar' }).length, 0, 'the Apply/Cancel row is gone; edits commit themselves');
    const notes = view.root.findAllByProps({ className: 'ribbon-note' }).map(text);
    assert.equal(notes.includes('Select text to format'), false, 'the tab row carries no formatting scope note');
  } finally {
    await act(async () => view.unmount());
  }
});

function twoParagraphDocument() {
  const model = tinyDocument('First');
  const second = tinyDocument('Second');
  const paragraph = second.body.blocks[0].paragraph;
  paragraph.id = 'p2'; paragraph.anchor.path = 'w:p[1]';
  paragraph.runs[0].id = 'r2'; paragraph.runs[0].anchor.path = 'w:p[1]/w:r[0]';
  model.body.blocks.push({ paragraph });
  model._targets.push({ ...second._targets[0], key: 'run-2', runId: 'r2', paragraphId: 'p2' });
  return model;
}

function clickRun(view, id) {
  view.root.findByProps({ 'data-docx-run': id }).props.onClick({ clientX: 20, clientY: 30 });
}
function typeDraft(view, text) {
  view.root.findByProps({ role: 'textbox' }).props.onInput({ currentTarget: { textContent: text } });
}

test('a rejected draft stays visible while another paragraph accepts the caret and typing', async t => {
  mockWindow();
  const client = mockClient(twoParagraphDocument());
  let calls = 0, commit, recovery;
  client.apply = async () => { calls++; throw new Error('Run is read-only'); };
  const { view } = await mountEditor(client, { registerCommit: fn => { commit = fn; }, onRecoveryDraftChange: value => { recovery = value; } });
  t.mock.timers.enable({ apis: ['setTimeout'] });
  try {
    await act(async () => clickRun(view, 'r1'));
    await act(async () => typeDraft(view, 'First draft'));
    await act(async () => documentPreview(view).props.commit());
    assert.equal(calls, 1);
    assert.match(view.root.findByProps({ role: 'alert' }).children.join(''), /read-only/);
    await act(async () => clickRun(view, 'r2'));
    assert.equal(documentPreview(view).props.selected, 'run-2');
    assert.equal(view.root.findByProps({ 'data-docx-run': 'r1' }).children.join(''), 'First draft');
    await act(async () => typeDraft(view, 'Second draft'));
    assert.equal(documentPreview(view).props.draft, 'Second draft');
    assert.deepEqual(recovery.drafts, { 'run-1': 'First draft', 'run-2': 'Second draft' });
    await act(async () => clickRun(view, 'r1'));
    assert.equal(calls, 2, 'only the newly edited second paragraph is attempted');
    assert.equal(documentPreview(view).props.draft, 'First draft');
    let saved;
    await act(async () => { saved = await commit(); });
    assert.equal(saved, false, 'uncommitted drafts must never be reported as saved');
    assert.equal(calls, 2, 'the same rejected value does not retry on clicks or Save');
  } finally { await act(async () => view.unmount()); }
});

test('a click during a rejecting commit is honoured after it settles', async t => {
  mockWindow();
  const client = mockClient(twoParagraphDocument());
  let reject, calls = 0;
  client.apply = () => { calls++; return new Promise((_, fail) => { reject = fail; }); };
  const { view, busy } = await mountEditor(client);
  t.mock.timers.enable({ apis: ['setTimeout'] });
  try {
    await act(async () => clickRun(view, 'r1'));
    await act(async () => typeDraft(view, 'First draft'));
    await act(async () => documentPreview(view).props.commit());
    assert.equal(busy.at(-1), true);
    assert.equal(view.root.findByProps({ role: 'textbox' }).props.contentEditable, 'plaintext-only');
    await act(async () => clickRun(view, 'r2'));
    await act(async () => reject(new Error('Run is read-only')));
    assert.equal(documentPreview(view).props.selected, 'run-2');
    assert.equal(calls, 1);
    await act(async () => typeDraft(view, 'Still usable'));
    assert.equal(documentPreview(view).props.draft, 'Still usable');
    assert.equal(view.root.findByProps({ 'data-docx-run': 'r1' }).children.join(''), 'First draft');
  } finally { await act(async () => view.unmount()); }
});

for (const continued of ['First draft continued', 'First']) test(`typing ${JSON.stringify(continued)} during a worker call survives readback and commits next`, async t => {
  mockWindow();
  let model = twoParagraphDocument(), resolve;
  const submitted = [];
  const client = { extract: async () => model, terminate() {}, apply: (bytes, _, envelope) => {
    const text = envelope.payload.mutations[0].text;
    submitted.push(text);
    return new Promise(done => { resolve = () => {
      model = structuredClone(model);
      model._targets[0].text = text;
      model.body.blocks[0].paragraph.runs[0].text = text;
      done(bytes);
    }; });
  } };
  const { view } = await mountEditor(client);
  t.mock.timers.enable({ apis: ['setTimeout'] });
  try {
    await act(async () => clickRun(view, 'r1'));
    await act(async () => typeDraft(view, 'First draft'));
    await act(async () => { t.mock.timers.tick(IDLE_COMMIT_MS); });
    await act(async () => typeDraft(view, continued));
    await act(async () => resolve());
    assert.equal(documentPreview(view).props.draft, continued);
    assert.equal(documentPreview(view).props.selected, 'run-1');
    await act(async () => { t.mock.timers.tick(IDLE_COMMIT_MS); });
    assert.deepEqual(submitted, ['First draft', continued]);
    await act(async () => resolve());
    assert.equal(documentPreview(view).props.hasDraft, false);
  } finally { await act(async () => view.unmount()); }
});

test('a successful edit in another paragraph retains failed-draft recovery', async t => {
  mockWindow();
  let model = twoParagraphDocument(), recovery;
  const client = { extract: async () => model, terminate() {}, apply: async (bytes, _, envelope) => {
    const mutation = envelope.payload.mutations[0];
    if (mutation.target_id === 'r1') throw new Error('Run is read-only');
    model = structuredClone(model);
    model._targets[1].text = mutation.text;
    model.body.blocks[1].paragraph.runs[0].text = mutation.text;
    return bytes;
  } };
  const { view } = await mountEditor(client, {
    onChange: () => { recovery = null; },
    onRecoveryDraftChange: value => { recovery = value; },
  });
  t.mock.timers.enable({ apis: ['setTimeout'] });
  try {
    await act(async () => clickRun(view, 'r1'));
    await act(async () => typeDraft(view, 'Keep this rejected draft'));
    await act(async () => documentPreview(view).props.commit());
    await act(async () => clickRun(view, 'r2'));
    await act(async () => typeDraft(view, 'Second committed'));
    await act(async () => documentPreview(view).props.commit());
    assert.equal(documentPreview(view).props.hasDraft, true);
    assert.deepEqual(recovery.drafts, { 'run-1': 'Keep this rejected draft' });
    assert.equal(view.root.findByProps({ 'data-docx-run': 'r1' }).children.join(''), 'Keep this rejected draft');
  } finally { await act(async () => view.unmount()); }
  const restored = await mountEditor(client, { initialRecoveryDraft: recovery });
  try {
    assert.equal(documentPreview(restored.view).props.draft, 'Keep this rejected draft');
    assert.equal(documentPreview(restored.view).props.hasDraft, true);
  } finally { await act(async () => restored.view.unmount()); }
});

test('reverting to saved text while a commit rejects clears the pending draft', async t => {
  mockWindow();
  let reject, commit;
  const client = mockClient(twoParagraphDocument());
  client.apply = () => new Promise((_, fail) => { reject = fail; });
  const { view } = await mountEditor(client, { registerCommit: fn => { commit = fn; } });
  t.mock.timers.enable({ apis: ['setTimeout'] });
  try {
    await act(async () => clickRun(view, 'r1'));
    await act(async () => typeDraft(view, 'First draft'));
    await act(async () => documentPreview(view).props.commit());
    await act(async () => typeDraft(view, 'First'));
    await act(async () => reject(new Error('Run is read-only')));
    assert.equal(documentPreview(view).props.draft, 'First');
    assert.equal(documentPreview(view).props.hasDraft, false);
    let saved;
    await act(async () => { saved = await commit(); });
    assert.equal(saved, true);
  } finally { await act(async () => view.unmount()); }
});

test('initial engine parse rejection is reported to the workspace', async () => {
  const errors = [];
  const Editor = await loadEditor({ extract: async () => { throw new Error('invalid XML in word/document.xml'); }, terminate() {} });
  mockWindow(); let renderer;
  try {
    await act(async () => { renderer = create(React.createElement(Editor, { name: 'Corrupt.docx', bytes: new Uint8Array([80, 75, 3, 4]), onChange() {}, onInitialLoadError: reason => errors.push(reason) })); });
    await until(() => errors.length > 0);
    assert.match(errors[0], /invalid XML in word\/document.xml/);
  } finally { if (renderer) await act(async () => renderer.unmount()); delete global.window; }
});
