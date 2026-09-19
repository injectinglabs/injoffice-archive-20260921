const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const React = require('react');
const { create, act } = require('react-test-renderer');
global.IS_REACT_ACT_ENVIRONMENT = true;

async function loadEditor() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/PresentationEditor.tsx'),
    platform: 'node',
    external: id => /^react(?:\/|$)/.test(id) || id === '@injoffice/pptx-native' || id.includes('playground/src/pptxRoundTrip'),
    transform: { jsx: { runtime: 'automatic' } },
    plugins: [{
      name: 'presentation-editor-test',
      resolveId(id) {
        if (id.endsWith('.css')) return '\0css';
        if (id === '@injoffice/pptx-wasm') return '\0wasm';
        if (id === '@injoffice/pptx-render') return '\0render';
        if (id === './PresentationPlayer' || id.endsWith('/PresentationPlayer') || id.endsWith('/PresentationPlayer.tsx')) return '\0player';
      },
      load(id) {
        if (id === '\0css') return 'export default ""';
        if (id === '\0wasm') return 'export const createPptxWasmClient = () => globalThis.__pptxClient;';
        if (id === '\0render') return 'export const exportNativePptxSlideSvg = async () => "<svg />";';
        if (id === '\0player') {
          return 'import React from "react";\n'
            + 'export default function PresentationPlayer(props) {\n'
            + '  globalThis.__presentationPlayer = props;\n'
            + '  return React.createElement("div", { "aria-label": "PresentationPlayer" });\n'
            + '}';
        }
      },
    }],
  });
  try {
    const { output } = await bundle.generate({ format: 'cjs', codeSplitting: false });
    const mod = { exports: {} };
    const req = id => (id.includes('pptxRoundTrip') ? { editablePptxTextTargets: () => [], editablePptxShapeTargets: () => [] } : require(id));
    new Function('require', 'module', 'exports', output[0].code)(req, mod, mod.exports);
    return mod.exports.default ?? mod.exports;
  } finally {
    await bundle.close();
  }
}

function mockClient() {
  const applied = [];
  const compatibility = { status: 'editable', diagnostics: [] };
  const size = { cx: 9144000, cy: 5143500 };
  let revision = 'rev-1';
  let slides = [{ id: 's1', elements: [], compatibility }];
  function deck() {
    return {
      contractVersion: 'pptx-native/v1',
      documentId: 'deck',
      sourceRevision: revision,
      size,
      assets: [],
      slides: slides.map(slide => ({ ...slide })),
      compatibility,
    };
  }
  return {
    applied,
    extract: async () => deck(),
    apply: async (_bytes, _deck, request) => {
      applied.push(request);
      if (request.operations[0]?.kind === 'slide.insert') {
        slides = [...slides, { id: `s${slides.length + 1}`, elements: [], compatibility }];
        revision = `rev-${slides.length}`;
      }
      return new Uint8Array([slides.length]);
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

function button(view, label) {
  return view.root.findAllByType('button').find(node => node.props.children === label);
}

test('PresentationEditor reports busy, applies a slide insert, and Present mounts PresentationPlayer', async () => {
  assert.equal(fs.existsSync(path.resolve(__dirname, '../src/presentation-editor.css')), true);
  const client = mockClient();
  globalThis.__pptxClient = client;
  const PresentationEditor = await loadEditor();
  const busy = [];
  const changes = [];
  let view;
  try {
    await act(async () => {
      view = create(React.createElement(PresentationEditor, {
        name: 'Deck.pptx',
        bytes: new Uint8Array([1]),
        onChange: value => changes.push(value),
        onBusyChange: value => busy.push(value),
      }));
    });
    await until(() => busy.at(-1) === false && view.root.findAllByProps({ 'aria-label': 'Show slide 1' }).length > 0);
    assert.equal(view.root.findByProps({ 'aria-label': 'Presentation editor' }) != null, true);
    assert.equal(busy.includes(true), true);
    assert.equal(busy.at(-1), false);
    await act(async () => button(view, 'New slide').props.onClick());
    await until(() => changes.length === 1 && busy.at(-1) === false && view.root.findAllByProps({ 'aria-label': 'Show slide 2' }).length > 0);
    assert.equal(client.applied[0].operations[0].kind, 'slide.insert');
    assert.deepEqual([...changes[0]], [2]);
    await act(async () => button(view, 'Present').props.onClick());
    assert.equal(globalThis.__presentationPlayer.initial.index, 1);
    assert.equal(globalThis.__presentationPlayer.initial.deck.slides.length, 2);
    assert.equal(view.root.findByProps({ 'aria-label': 'PresentationPlayer' }) != null, true);
  } finally {
    if (view) await act(async () => view.unmount());
    delete globalThis.__pptxClient;
    delete globalThis.__presentationPlayer;
  }
});
