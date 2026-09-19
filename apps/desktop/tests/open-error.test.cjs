const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const React = require('react');
const { create, act } = require('react-test-renderer');
global.IS_REACT_ACT_ENVIRONMENT = true;

async function loadOpenError() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/OpenError.tsx'),
    platform: 'node',
    external: id => /^react(?:\/|$)/.test(id),
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
    new Function('require', 'module', 'exports', output[0].code)(require, mod, mod.exports);
    return mod.exports.default ?? mod.exports;
  } finally {
    await bundle.close();
  }
}

test('open error explains unsupported files and offers Open another or Back to start', async () => {
  const OpenError = await loadOpenError();
  const opened = [];
  let home = 0;
  let view;
  await act(async () => {
    view = create(React.createElement(OpenError, {
      name: 'report.doc',
      kind: 'unsupported',
      onOpen: () => opened.push('open'),
      onHome: () => { home += 1; },
    }));
  });
  assert.match(view.root.findByProps({ id: 'open-error-title' }).children.join(''), /not supported/i);
  assert.equal(view.root.findByProps({ className: 'open-error-name' }).children.join(''), 'report.doc');
  await act(async () => view.root.findByProps({ className: 'open-error-open' }).props.onClick());
  assert.deepEqual(opened, ['open']);
  await act(async () => view.root.findAllByType('button').find(button => button.props.children === 'Back to start').props.onClick());
  assert.equal(home, 1);

  await act(async () => {
    view.update(React.createElement(OpenError, {
      kind: 'encrypted',
      busy: true,
      onOpen() {},
      onHome() {},
    }));
  });
  assert.match(view.root.findByProps({ id: 'open-error-title' }).children.join(''), /encrypted/i);
  assert.equal(view.root.findByProps({ className: 'open-error-open' }).props.disabled, true);
});
