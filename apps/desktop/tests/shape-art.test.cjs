const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const React = require('react');
const { create, act } = require('react-test-renderer');
global.IS_REACT_ACT_ENVIRONMENT = true;

async function loadShapeArt() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/ShapeArt.tsx'),
    platform: 'node',
    external: id => /^react(?:\/|$)/.test(id) || id === '@injoffice/pptx-native',
    transform: { jsx: { runtime: 'automatic' } },
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

test('ShapeArt maps ellipse, triangle, and roundRect presets to SVG primitives', async () => {
  const ShapeArt = await loadShapeArt();
  const props = { width: 100, height: 50, fill: '#fff', stroke: '#000', strokeWidth: 1 };
  let view;
  await act(async () => { view = create(React.createElement(ShapeArt, { ...props, element: { kind: 'shape', preset: 'ellipse' } })); });
  assert.equal(view.root.findAllByType('ellipse').length, 1);
  await act(async () => { view.update(React.createElement(ShapeArt, { ...props, element: { kind: 'shape', preset: 'triangle' } })); });
  assert.equal(view.root.findAllByType('polygon').length, 1);
  await act(async () => { view.update(React.createElement(ShapeArt, { ...props, element: { kind: 'shape', preset: 'roundRect' } })); });
  assert.equal(view.root.findByType('rect').props.rx, 6);
  await act(async () => { view.update(React.createElement(ShapeArt, { ...props, element: { kind: 'connector', flipH: true, stroke: { cap: 'flat' } } })); });
  assert.equal(view.root.findByType('line').props.y1, 50);
});
