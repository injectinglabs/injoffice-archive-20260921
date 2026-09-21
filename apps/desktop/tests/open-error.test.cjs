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

test('classifyOpenError maps host messages onto unsupported, encrypted, and extract pages', async () => {
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
    const { classifyOpenError, containerFailure } = mod.exports;
    assert.equal(classifyOpenError('Choose a DOCX document, XLSX workbook, PPTX presentation, or PDF.'), 'unsupported');
    assert.equal(classifyOpenError('This document is encrypted.'), 'encrypted');
    assert.equal(classifyOpenError('Native extract refused the package.'), 'extract');
    assert.equal(classifyOpenError('Unknown document. Open the file again.'), undefined);
    // Office Open XML files are ZIP packages; a renamed text file is refused before any engine runs.
    const zip = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0]);
    const text = new Uint8Array([0x54, 0x68, 0x69, 0x73]);
    for (const extension of ['docx', 'xlsx', 'pptx']) {
      assert.equal(containerFailure(`Report.${extension}`, zip), undefined);
      assert.match(containerFailure(`Report.${extension}`, text), /does not start with a ZIP package \(PK\): the first bytes are 54 68 69 73/);
    }
    assert.equal(containerFailure('Scan.pdf', new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d])), undefined);
    assert.match(containerFailure('Scan.pdf', zip), /does not start with %PDF/);
    assert.match(containerFailure('Report.docx', new Uint8Array()), /is empty \(0 bytes\)/);
    assert.equal(containerFailure('notes.txt', text), undefined, 'other extensions are the host dialog\'s business');
  } finally {
    await bundle.close();
  }
});

test('open error names the file, keeps the engine string in Details, and offers Open another or the start page', async () => {
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
  assert.match(view.root.findByProps({ id: 'open-error-title' }).children.join(''), /InjOffice can.t open this file/i);
  assert.equal(view.root.findByProps({ className: 'open-error-name' }).children.join(''), 'report.doc');
  assert.match(view.root.findByProps({ className: 'open-error-reason' }).children.join(''), /opens \.docx, \.xlsx, \.pptx and \.pdf/);
  assert.equal(view.root.findAllByProps({ className: 'open-error-details' }).length, 0, 'no Details section without an engine string');
  await act(async () => view.root.findByProps({ className: 'open-error-open' }).props.onClick());
  assert.deepEqual(opened, ['open']);
  await act(async () => view.root.findAllByType('button').find(button => button.props.children === 'Go to start page').props.onClick());
  assert.equal(home, 1);

  await act(async () => {
    view.update(React.createElement(OpenError, {
      kind: 'encrypted',
      detail: 'docxpatch: native extract: invalid ZIP: zip: not a valid zip file',
      busy: true,
      onOpen() {},
      onHome() {},
    }));
  });
  assert.match(view.root.findByProps({ id: 'open-error-title' }).children.join(''), /InjOffice can.t open this file/i);
  assert.match(view.root.findByProps({ className: 'open-error-reason' }).children.join(''), /password-protected/i);
  // The raw engine string is available but folded away, the way Office keeps its details.
  const details = view.root.findByProps({ className: 'open-error-details' });
  assert.equal(details.findByType('summary').children.join(''), 'Details');
  assert.match(details.findByType('p').children.join(''), /invalid ZIP/);
  assert.equal(view.root.findByProps({ className: 'open-error-open' }).props.disabled, true);
});
