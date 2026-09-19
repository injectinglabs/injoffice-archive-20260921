const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const React = require('react');
const { create, act } = require('react-test-renderer');
global.IS_REACT_ACT_ENVIRONMENT = true;

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
    return mod.exports.default ?? mod.exports;
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
  await act(async () => textbox.props.onInput({ currentTarget: { textContent: 'Hello' } }));
  const apply = view.root.findByProps({ className: 'office-apply' });
  assert.equal(apply.props.disabled, false);
  await act(async () => apply.props.onClick());
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
  await act(async () => { t.mock.timers.tick(79); });
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
  await act(async () => { t.mock.timers.tick(80); });
  for (let count = 0; count < 10; count++) await act(async () => Promise.resolve());
  assert.equal(client.applied.length, 1);
  await act(async () => view.unmount());
});
