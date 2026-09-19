const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const React = require('react');
const { create, act } = require('react-test-renderer');
global.IS_REACT_ACT_ENVIRONMENT = true;

async function loadMode() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/presentationMode.ts'),
    platform: 'node',
    external: id => id === '@injoffice/pptx-native',
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

async function loadPlayer() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/PresentationPlayer.tsx'),
    platform: 'node',
    external: id => /^react(?:\/|$)/.test(id) || id === '@injoffice/pptx-native',
    transform: { jsx: { runtime: 'automatic' } },
    plugins: [{
      name: 'presentation-player-test',
      resolveId(id) {
        if (id.endsWith('.css')) return '\0css';
        if (id === './presentationMode' || id.endsWith('/presentationMode') || id.endsWith('/presentationMode.ts')) return '\0mode';
      },
      load(id) {
        if (id === '\0css') return 'export default ""';
        if (id === '\0mode') {
          return 'export function navigatePresentation(state, key) { return globalThis.__navigatePresentation(state, key); }\n'
            + 'export function presentationScale(...args) { return globalThis.__presentationScale(...args); }';
        }
      },
    }],
  });
  try {
    const { output } = await bundle.generate({ format: 'cjs', codeSplitting: false });
    const mod = { exports: {} };
    new Function('require', 'module', 'exports', output[0].code)(require, mod, mod.exports);
    return mod.exports.default ?? mod.exports;
  } finally {
    await bundle.close();
  }
}

function hostRefs() {
  const original = React.useRef;
  React.useRef = function (initial) {
    if (initial != null) return original.call(this, initial);
    const host = { showModal() {}, close() {}, focus() {}, requestFullscreen: async () => {} };
    const ref = original.call(this, host);
    return new Proxy(ref, {
      set(obj, prop, value) {
        if (prop === 'current' && value == null) return true;
        obj[prop] = value;
        return true;
      },
    });
  };
  return () => { React.useRef = original; };
}

function installDom() {
  global.HTMLInputElement = class HTMLInputElement {};
  global.HTMLButtonElement = class HTMLButtonElement {};
  global.ResizeObserver = class {
    constructor(callback) { this.callback = callback; }
    observe() { this.callback([{ contentRect: { width: 1920, height: 1080 } }]); }
    disconnect() {}
  };
  global.document = {
    activeElement: null,
    fullscreenElement: null,
    addEventListener() {},
    removeEventListener() {},
    exitFullscreen: async () => {},
  };
}

function slide(id) {
  return { id, compatibility: { status: 'editable', diagnostics: [] } };
}

function deck() {
  return { size: { cx: 9144000, cy: 5143500 }, slides: [slide('a'), slide('b'), slide('c')] };
}

function statusText(view) {
  return view.root.findByProps({ 'aria-label': 'Presentation navigation' }).findByProps({ role: 'status' }).children.join('');
}

function press(view, key) {
  view.root.findByProps({ className: 'presentation-player' }).props.onKeyDown({
    key, preventDefault() {}, stopPropagation() {},
  });
}

test('PresentationPlayer Escape calls onExit; Home/Arrow navigation uses presentationMode helpers', async () => {
  const mode = await loadMode();
  const PresentationPlayer = await loadPlayer();
  const navigated = [];
  globalThis.__navigatePresentation = (state, key) => {
    navigated.push(key);
    return mode.navigatePresentation(state, key);
  };
  globalThis.__presentationScale = (...args) => mode.presentationScale(...args);
  installDom();
  const restoreRefs = hostRefs();
  const exits = [];
  const shown = [];
  const initial = mode.startPresentationMode(deck(), 1, false);
  let view;
  try {
    await act(async () => {
      view = create(React.createElement(PresentationPlayer, {
        initial,
        renderSlide: slide => { shown.push(slide.id); return slide.id; },
        onExit: () => exits.push('exit'),
      }));
    });
    assert.equal(statusText(view), '2 / 3');
    assert.equal(shown.at(-1), 'b');
    await act(async () => press(view, 'Home'));
    assert.deepEqual(navigated, ['Home']);
    assert.equal(statusText(view), '1 / 3');
    assert.equal(shown.at(-1), 'a');
    await act(async () => press(view, 'ArrowRight'));
    assert.deepEqual(navigated, ['Home', 'ArrowRight']);
    assert.equal(statusText(view), '2 / 3');
    await act(async () => press(view, 'Escape'));
    assert.deepEqual(exits, ['exit']);
    await act(async () => press(view, 'Escape'));
    assert.deepEqual(exits, ['exit']);
  } finally {
    if (view) await act(async () => view.unmount());
    restoreRefs();
    delete global.document;
    delete global.ResizeObserver;
    delete global.HTMLInputElement;
    delete global.HTMLButtonElement;
    delete globalThis.__navigatePresentation;
    delete globalThis.__presentationScale;
  }
});
