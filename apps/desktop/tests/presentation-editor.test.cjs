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
    const req = id => (id.includes('pptxRoundTrip') ? { editablePptxTextTargets: deck => globalThis.__textTargets?.(deck) ?? [], editablePptxShapeTargets: () => [] } : require(id));
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
  const source = { partName: 'ppt/slides/slide1.xml', objectId: 'slide', fingerprintSha256: 'a'.repeat(64) };
  let slides = [{ id: 's1', elements: [], compatibility, source }];
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
        slides = [...slides, { id: `s${slides.length + 1}`, elements: [], compatibility, source }];
        revision = `rev-${slides.length}`;
      }
      if (request.operations[0]?.kind === 'slide.background.set') {
        const op = request.operations[0];
        slides = slides.map(slide => slide.id === op.slideId ? { ...slide, background: op.fill } : slide);
        revision += '-background';
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

const text = node => node.children.map(child => typeof child === 'string' ? child : text(child)).join('');
function button(view, label) {
  return view.root.findAllByType('button').find(node => node.props.children === label || text(node) === label);
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
    assert.equal(client.applied[0].operations[0].slideId, 's1');
    assert.equal(client.applied[0].operations[0].expectedFingerprintSha256, 'a'.repeat(64));
    assert.equal(Object.hasOwn(client.applied[0].operations[0], 'elementId'), false);
    await act(async () => view.root.findByProps({ 'aria-label': 'Accent background' }).props.onClick());
    await until(() => changes.length === 2 && busy.at(-1) === false);
    assert.equal(client.applied[1].operations[0].kind, 'slide.background.set');
    assert.equal(client.applied[1].operations[0].slideId, 's2');
    assert.equal(client.applied[1].operations[0].fill, '2459AD');
    assert.equal(view.root.findByProps({ 'aria-label': 'Slide canvas' }).props.style.background, '#2459AD');
    assert.deepEqual([...changes[0]], [2]);
    await act(async () => button(view, 'From Beginning').props.onClick());
    assert.equal(globalThis.__presentationPlayer.initial.index, 0, 'From Beginning starts at the first slide');
    assert.equal(globalThis.__presentationPlayer.initial.deck.slides.length, 2);
    assert.equal(view.root.findByProps({ 'aria-label': 'PresentationPlayer' }) != null, true);
  } finally {
    if (view) await act(async () => view.unmount());
    delete globalThis.__pptxClient;
    delete globalThis.__presentationPlayer;
  }
});

test('PresentationEditor arranges its controls as a PowerPoint ribbon with labelled groups and icons', async () => {
  const client = mockClient();
  globalThis.__pptxClient = client;
  const PresentationEditor = await loadEditor();
  const busy = [];
  let view;
  try {
    await act(async () => {
      view = create(React.createElement(PresentationEditor, { name: 'Deck.pptx', bytes: new Uint8Array([1]), onChange: () => {}, onBusyChange: value => busy.push(value) }));
    });
    await until(() => busy.at(-1) === false && view.root.findAllByProps({ 'aria-label': 'Show slide 1' }).length > 0);
    assert.deepEqual(view.root.findAllByProps({ role: 'tab' }).map(text), ['File', 'Home', 'Insert', 'Design', 'Slide Show']);
    const panels = view.root.findAllByProps({ role: 'tabpanel' });
    const groups = panel => panel.findAllByProps({ role: 'group' }).map(group => group.props['aria-label']);
    assert.deepEqual(groups(panels[0]), ['Export']);
    assert.deepEqual(groups(panels[1]), ['Slides', 'Font', 'Paragraph', 'Drawing']);
    assert.deepEqual(groups(panels[2]), ['Tables', 'Images', 'Illustrations', 'Text']);
    assert.deepEqual(groups(panels[3]), ['Customize']);
    assert.deepEqual(groups(panels[4]), ['Start Slide Show']);
    assert.equal(view.root.findAllByProps({ role: 'toolbar' }).filter(toolbar => toolbar.props['aria-label'] === 'Quick access').length, 0, 'Undo/Redo live in the shell title bar, not in the ribbon');
    const commandButtons = view.root.findAllByProps({ role: 'group' }).flatMap(group => group.findAllByType('button'));
    assert.ok(commandButtons.length >= 16, `command buttons: ${commandButtons.length}`);
    for (const node of commandButtons) {
      assert.equal(node.findAllByType('svg').length, 1, `${node.props.title} has an icon`);
      assert.ok(node.props.title, 'every command button has a tooltip');
    }
    const byLabel = Object.fromEntries(commandButtons.map(node => [node.props['aria-label'] ?? text(node), node]));
    assert.match(byLabel.Underline.props.title, /not supported by the native PPTX transaction/);
    assert.match(byLabel.Bullets.props.title, /not supported by the native PPTX transaction/);
    assert.equal(byLabel['From Beginning'].props.disabled, false);
    assert.equal(byLabel['From Current Slide'].props.disabled, false);
    const gallery = view.root.findByProps({ 'aria-label': 'Slide background' });
    assert.equal(gallery.findAllByType('select').length, 0, 'the background is a gallery, not a select');
    assert.deepEqual(gallery.findAllByType('button').map(node => node.props['aria-label']), ['White background', 'Light grey background', 'Dark background', 'Accent background']);
  } finally {
    if (view) await act(async () => view.unmount());
    delete globalThis.__pptxClient;
  }
});

test('slide thumbnails derive their width from the rail: number on the left, deck aspect ratio, no fixed frame width', async () => {
  const client = mockClient();
  globalThis.__pptxClient = client;
  const PresentationEditor = await loadEditor();
  const busy = [];
  let view;
  try {
    await act(async () => {
      view = create(React.createElement(PresentationEditor, { name: 'Deck.pptx', bytes: new Uint8Array([1]), onChange: () => {}, onBusyChange: value => busy.push(value) }));
    });
    await until(() => busy.at(-1) === false && view.root.findAllByProps({ 'aria-label': 'Show slide 1' }).length > 0);
    const thumbnail = view.root.findByProps({ 'aria-label': 'Show slide 1' });
    assert.equal(thumbnail.props.className, 'presentation-thumbnail');
    const [number, stage] = thumbnail.children;
    assert.equal(number.props.className, 'presentation-slide-number', 'the slide number comes first, to the left of the thumbnail');
    assert.equal(number.children.join(''), '1');
    assert.equal(stage.props.className, 'presentation-thumb-stage');
    assert.equal(stage.props.style.aspectRatio, '9144000 / 5143500', 'the stage keeps the deck aspect ratio');
    assert.equal(stage.props.style.width, undefined, 'the stage takes its width from the pane, not from a fixed number');
    const css = fs.readFileSync(path.resolve(__dirname, '../src/presentation-editor.css'), 'utf8');
    assert.match(css, /\.presentation-thumb-stage \{[^}]*width: 100%/, 'the stage fills the rail column');
    assert.match(css, /\.presentation-thumb-stage \{[^}]*overflow: hidden/);
    const source = fs.readFileSync(path.resolve(__dirname, '../src/PresentationEditor.tsx'), 'utf8');
    assert.equal(/148 \/ \(/.test(source), false, 'no hard-coded 148px thumbnail width');
  } finally {
    if (view) await act(async () => view.unmount());
    delete globalThis.__pptxClient;
  }
});

test('the Format pane is closed until Format is pressed and Arrange is one of its groups', async () => {
  const client = mockClient();
  globalThis.__pptxClient = client;
  const PresentationEditor = await loadEditor();
  const busy = [];
  let view;
  try {
    await act(async () => {
      view = create(React.createElement(PresentationEditor, { name: 'Deck.pptx', bytes: new Uint8Array([1]), onChange: () => {}, onBusyChange: value => busy.push(value) }));
    });
    await until(() => busy.at(-1) === false && view.root.findAllByProps({ 'aria-label': 'Show slide 1' }).length > 0);
    const pane = () => view.root.findAllByType('aside').filter(node => node.props['aria-label'] === 'Format');
    assert.equal(pane().length, 0, 'nothing is selected, so the slide keeps the width');
    const format = view.root.findAllByType('button').find(node => text(node) === 'Format');
    assert.equal(format.props['aria-pressed'], false);
    await act(async () => format.props.onClick());
    assert.equal(pane().length, 1, 'Format opens the pane on demand');
    assert.equal(view.root.findAllByType('button').find(node => text(node) === 'Format').props['aria-pressed'], true);
    await act(async () => view.root.findByProps({ 'aria-label': 'Close the Format pane' }).props.onClick());
    assert.equal(pane().length, 0);
    const arrange = view.root.findAllByType('button').find(node => text(node) === 'Arrange');
    assert.equal(arrange.props.disabled, true, 'this slide has fewer than two arrangeable objects');
    const source = fs.readFileSync(path.resolve(__dirname, '../src/SlideArrangePanel.tsx'), 'utf8');
    assert.equal(/<details|<summary/.test(source), false, 'Arrange is a pane group, not a disclosure link');
    assert.match(source, /aria-label="Arrange objects"/);
  } finally {
    if (view) await act(async () => view.unmount());
    delete globalThis.__pptxClient;
  }
});

/** A deck with one editable text box, so the caret can be put in painted text. */
function textClient() {
  const applied = [];
  const compatibility = { status: 'editable', diagnostics: [] };
  const paragraphs = () => [{ align: 'left', runs: [{ text: 'Hello', fontFamily: 'Arial', fontSizeHundredthPt: 1800, bold: false, italic: false, color: '20242B' }] }];
  let text = 'Hello';
  let revision = 'rev-1';
  const element = () => ({ id: 'e1', kind: 'text', name: 'Title', source: { partName: 'ppt/slides/slide1.xml', objectId: '2' }, compatibility, rotation60000: 0,
    transform: { x: 0, y: 0, cx: 4572000, cy: 1143000 }, textBody: { leftInsetEmu: 0, rightInsetEmu: 0, topInsetEmu: 0, bottomInsetEmu: 0, wrap: 'square', verticalAnchor: 'top' },
    paragraphs: [{ ...paragraphs()[0], runs: [{ ...paragraphs()[0].runs[0], text }] }] });
  const deck = () => ({ contractVersion: 'pptx-native/v1', documentId: 'deck', sourceRevision: revision, size: { cx: 9144000, cy: 5143500 }, assets: [], compatibility,
    slides: [{ id: 's1', elements: [element()], compatibility }] });
  globalThis.__textTargets = value => value.slides[0].elements.map(item => ({ slideIndex: 0, elementId: item.id, sourcePartName: item.source.partName, sourceObjectId: item.source.objectId, expectedFingerprintSha256: 'f', paragraphs: item.paragraphs }));
  return {
    applied,
    extract: async () => deck(),
    apply: async (_bytes, _deck, request) => { applied.push(request); text = request.operations[0].paragraphs[0].runs[0].text; revision = `rev-${applied.length + 1}`; return new Uint8Array([1]); },
    terminate() {},
  };
}

test('clicking slide text puts the caret on the canvas; Enter commits through the text transaction', async () => {
  const client = textClient();
  globalThis.__pptxClient = client;
  const PresentationEditor = await loadEditor();
  const busy = [];
  let view;
  try {
    await act(async () => {
      view = create(React.createElement(PresentationEditor, { name: 'Deck.pptx', bytes: new Uint8Array([1]), onChange: () => {}, onBusyChange: value => busy.push(value) }));
    });
    await until(() => busy.at(-1) === false && view.root.findAllByProps({ 'aria-label': 'Select Title' }).length > 0);
    assert.equal(view.root.findAllByProps({ role: 'textbox' }).length, 0, 'no caret before the text is clicked');
    const object = view.root.findByProps({ 'aria-label': 'Select Title' });
    await act(async () => object.props.onClick({ stopPropagation() {}, target: {}, clientX: 40, clientY: 40, ctrlKey: false, metaKey: false }));
    const caret = view.root.findByProps({ role: 'textbox' });
    assert.equal(caret.props['aria-label'], 'Slide text');
    assert.equal(caret.props.contentEditable, 'plaintext-only');
    assert.equal(caret.props.style.fontFamily, 'Arial', 'the caret keeps the painted font');
    await act(async () => caret.props.onInput({ currentTarget: { textContent: 'Hello world' } }));
    assert.equal(client.applied.length, 0, 'typing stays a draft');
    await act(async () => caret.props.onKeyDown({ key: 'Enter', preventDefault() {}, stopPropagation() {} }));
    await until(() => client.applied.length === 1 && busy.at(-1) === false);
    assert.equal(client.applied[0].operations[0].kind, 'text.replace');
    assert.equal(client.applied[0].operations[0].paragraphs[0].runs[0].text, 'Hello world');
    assert.equal(view.root.findAllByProps({ role: 'textbox' }).length, 0, 'the caret leaves on commit');
  } finally {
    if (view) await act(async () => view.unmount());
    delete globalThis.__pptxClient;
    delete globalThis.__textTargets;
  }
});

test('the inspector has no text segment select or textarea, and the wrapping caveat is a status-row tooltip', async () => {
  const client = textClient();
  globalThis.__pptxClient = client;
  const PresentationEditor = await loadEditor();
  const busy = [];
  let view;
  try {
    await act(async () => {
      view = create(React.createElement(PresentationEditor, { name: 'Deck.pptx', bytes: new Uint8Array([1]), onChange: () => {}, onBusyChange: value => busy.push(value) }));
    });
    await until(() => busy.at(-1) === false && view.root.findAllByProps({ 'aria-label': 'Select Title' }).length > 0);
    const object = view.root.findByProps({ 'aria-label': 'Select Title' });
    await act(async () => object.props.onClick({ stopPropagation() {}, target: {}, clientX: 40, clientY: 40, ctrlKey: false, metaKey: false }));
    assert.equal(view.root.findAllByProps({ 'aria-label': 'Slide text segment' }).length, 0);
    assert.equal(view.root.findAllByType('textarea').length, 0);
    assert.equal(view.root.findAllByType('button').filter(node => text(node) === 'Text' && node.props['aria-pressed'] !== undefined).length, 0, 'no Text tab in the pane');
    const note = view.root.findByProps({ 'aria-label': 'Preview note' });
    assert.match(note.props.title, /text wrapping may differ in PowerPoint/);
    assert.equal(view.root.findByProps({ 'aria-label': 'Presentation status' }) != null, true);
    // Esc puts the painted text back and drops the draft.
    const caret = view.root.findByProps({ role: 'textbox' });
    await act(async () => caret.props.onInput({ currentTarget: { textContent: 'Draft' } }));
    await act(async () => caret.props.onKeyDown({ key: 'Escape', preventDefault() {}, stopPropagation() {} }));
    assert.equal(client.applied.length, 0);
    assert.equal(view.root.findAllByProps({ role: 'textbox' }).length, 0);
  } finally {
    if (view) await act(async () => view.unmount());
    delete globalThis.__pptxClient;
    delete globalThis.__textTargets;
  }
});

test('the slide number lives in the status row; the canvas carries no caption or preview disclosure', async () => {
  const client = textClient();
  globalThis.__pptxClient = client;
  const PresentationEditor = await loadEditor();
  const busy = [];
  let view;
  try {
    await act(async () => {
      view = create(React.createElement(PresentationEditor, { name: 'Deck.pptx', bytes: new Uint8Array([1]), onChange: () => {}, onBusyChange: value => busy.push(value) }));
    });
    await until(() => busy.at(-1) === false && view.root.findAllByProps({ 'aria-label': 'Show slide 1' }).length > 0);
    assert.equal(view.root.findAllByType('details').length, 0, 'no Preview and editing limits disclosure');
    assert.equal(view.root.findAllByProps({ className: 'presentation-canvas-label' }).length, 0);
    const status = view.root.findByProps({ 'aria-label': 'Presentation status' });
    assert.match(text(status), /Slide 1 of 1/);
    const note = view.root.findByProps({ 'aria-label': 'Preview note' });
    assert.match(note.props.title, /text wrapping may differ in PowerPoint/);
    assert.match(note.props.title, /Unsupported content stays in the file/, 'the editing limits moved into the tooltip');
    const source = fs.readFileSync(path.resolve(__dirname, '../src/PresentationEditor.tsx'), 'utf8');
    assert.equal(/Positioned preview ·/.test(source), false);
  } finally {
    if (view) await act(async () => view.unmount());
    delete globalThis.__pptxClient;
    delete globalThis.__textTargets;
  }
});


test('PresentationEditor reports the first parser rejection to the workspace', async () => {
  const errors = [];
  globalThis.__pptxClient = { extract: async () => { throw new Error('invalid XML in pptx package'); }, terminate() {} };
  const Editor = await loadEditor();
  let view;
  try {
    await act(async () => { view = create(React.createElement(Editor, { name: 'Corrupt.pptx', bytes: new Uint8Array([80,75,3,4]), onChange() {}, onInitialLoadError: reason => errors.push(reason) })); });
    await until(() => errors.length > 0);
    assert.deepEqual(errors, ['invalid XML in pptx package']);
  } finally { if (view) await act(async () => view.unmount()); delete globalThis.__pptxClient; }
});


test('a presentation mutation rejection after load stays in the editor', async () => {
  const client = mockClient(), errors = [];
  client.apply = async () => { throw new Error('slide mutation refused'); };
  globalThis.__pptxClient = client;
  const Editor = await loadEditor();
  let view;
  try {
    await act(async () => { view = create(React.createElement(Editor, { name: 'Deck.pptx', bytes: new Uint8Array([1]), onChange() {}, onInitialLoadError: reason => errors.push(reason) })); });
    await until(() => view.root.findAllByProps({ 'aria-label': 'Show slide 1' }).length > 0);
    await act(async () => button(view, 'New slide').props.onClick());
    await until(() => JSON.stringify(view.toJSON()).includes('slide mutation refused'));
    assert.deepEqual(errors, []);
  } finally { if (view) await act(async () => view.unmount()); delete globalThis.__pptxClient; }
});
