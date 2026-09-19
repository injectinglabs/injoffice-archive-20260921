const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const React = require('react');
const { create, act } = require('react-test-renderer');
global.IS_REACT_ACT_ENVIRONMENT = true;

async function loadPanel() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/SlideArrangePanel.tsx'),
    platform: 'node',
    external: id => /^react(?:\/|$)/.test(id) || id === '@injoffice/pptx-native' || id === '@injoffice/pptx-wasm' || id.includes('playground/src/pptxRoundTrip'),
    transform: { jsx: { runtime: 'automatic' } },
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

test('arrange panel toggles objects and disables align until two are selected', async () => {
  const SlideArrangePanel = await loadPanel();
  const toggled = [];
  const arranged = [];
  const elements = [
    { id: 'a', kind: 'shape', name: 'Box', source: { partName: 'p', objectId: '1' }, compatibility: { status: 'editable' }, transform: { x: 0, y: 0, cx: 1, cy: 1 } },
    { id: 'b', kind: 'shape', name: 'Oval', source: { partName: 'p', objectId: '2' }, compatibility: { status: 'editable' }, transform: { x: 0, y: 0, cx: 1, cy: 1 } },
  ];
  let view;
  await act(async () => {
    view = create(React.createElement(SlideArrangePanel, {
      elements, keys: [], disabled: false,
      onToggle: key => toggled.push(key),
      onArrange: action => arranged.push(action),
    }));
  });
  assert.equal(view.root.findAllByType('button').find(button => button.props.children === 'Align left').props.disabled, true);
  await act(async () => view.root.findAllByType('input')[0].props.onChange());
  assert.equal(toggled.length, 1);
  await act(async () => {
    view.update(React.createElement(SlideArrangePanel, {
      elements, keys: [['p', '1'].join('\0'), ['p', '2'].join('\0')], disabled: false,
      onToggle: key => toggled.push(key),
      onArrange: action => arranged.push(action),
    }));
  });
  const left = view.root.findAllByType('button').find(button => button.props.children === 'Align left');
  assert.equal(left.props.disabled, false);
  await act(async () => left.props.onClick());
  assert.deepEqual(arranged, ['left']);
});
