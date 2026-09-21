const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const React = require('react');
const { create, act } = require('react-test-renderer');
global.IS_REACT_ACT_ENVIRONMENT = true;

function isExternal(id) {
  return /^react(?:\/|$)/.test(id) || id === 'pdf-lib' || id === '@pdf-lib/fontkit';
}

function stubRequire(id) {
  if (id === 'pdf-lib') {
    return {
      PDFArray: class PDFArray {},
      PDFDict: class PDFDict {},
      PDFDocument: class PDFDocument {},
      PDFHexString: { fromText() { return {}; } },
      PDFName: { of(name) { return { toString() { return '/' + name; } }; } },
      PDFTextField: class PDFTextField {},
      PDFCheckBox: class PDFCheckBox {},
      PDFDropdown: class PDFDropdown {},
      PDFOptionList: class PDFOptionList {},
      PDFRadioGroup: class PDFRadioGroup {},
      StandardFonts: { Helvetica: 'Helvetica' },
      degrees(angle) { return angle; },
      rgb(r, g, b) { return { r, g, b }; },
    };
  }
  if (id === '@pdf-lib/fontkit') {
    const fontkit = { create() { return {}; } };
    return { __esModule: true, default: fontkit, ...fontkit };
  }
  return require(id);
}

async function loadEditor() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/PdfEditor.tsx'),
    platform: 'node',
    external: isExternal,
    transform: { jsx: { runtime: 'automatic' } },
    plugins: [{
      name: 'pdf-editor-test',
      resolveId(id) {
        if (id.endsWith('.css') || id.includes('pdf.worker.min.mjs?url')) return '\0empty';
        if (id === 'pdfjs-dist/legacy/build/pdf.mjs') return '\0pdfjs';
        if (id === './pdf-commands' || id.endsWith('/pdf-commands') || id.endsWith('/pdf-commands.ts')) return '\0pdf-commands';
      },
      load(id) {
        if (id === '\0empty') return 'export default "worker.js"';
        if (id === '\0pdfjs') {
          return 'export const GlobalWorkerOptions={};'
            + 'export const getDocument=(options)=>globalThis.__pdfView.getDocument(options);'
            + 'export class TextLayer {textDivs=[];textContentItemsStr=["Page text"];async render(){}cancel(){}}';
        }
        if (id === '\0pdf-commands') {
          return [
            'export class PdfHistory {',
            '  constructor(bytes) { this.entries = [bytes.slice()]; this.position = 0; }',
            '  get bytes() { return this.entries[this.position]; }',
            '  get canUndo() { return this.position > 0; }',
            '  get canRedo() { return this.position + 1 < this.entries.length; }',
            '  push(bytes) { this.entries = [...this.entries.slice(0, this.position + 1), bytes.slice()]; this.position = this.entries.length - 1; return this.bytes; }',
            '  undo() { if (this.canUndo) this.position--; return this.bytes; }',
            '  redo() { if (this.canRedo) this.position++; return this.bytes; }',
            '}',
            'export function findPdfTextMatches() { return []; }',
            'export function parsePdfPageRange() { return [1]; }',
            'export async function inspectPdf(bytes) { return globalThis.__inspectPdf(bytes); }',
            'export async function applyPdfCommand(bytes, command) { return globalThis.__applyPdfCommand(bytes, command); }',
            'export async function importPdfPages() { throw new Error("import unused"); }',
            'export async function exportPdfPages() { throw new Error("export unused"); }',
          ].join('\n');
        }
      },
    }],
  });
  try {
    const { output } = await bundle.generate({ format: 'cjs', codeSplitting: false });
    const module = { exports: {} };
    new Function('require', 'module', 'exports', output[0].code)(stubRequire, module, module.exports);
    return module.exports.default ?? module.exports;
  } finally {
    await bundle.close();
  }
}

function summary(pages) {
  return {
    fields: [],
    pages: Array.from({ length: pages }, () => ({ width: 400, height: 500, rotation: 0, annotations: [] })),
    editable: true,
  };
}

async function until(predicate) {
  for (let i = 0; i < 100; i++) {
    if (predicate()) return;
    await act(async () => new Promise(resolve => setTimeout(resolve, 5)));
  }
  throw new Error('PDF editor did not settle');
}

function nodeMock(element) {
  if (element.type === 'canvas') return { getContext() { return {}; } };
  if (element.props.className === 'pdf-text-layer') return { replaceChildren() {} };
  return null;
}

/** A ribbon command button, located by its visible label (RibbonButton renders it in a span). */
function ribbonButton(renderer, label) {
  const found = renderer.root.findAllByType('button').filter(button => {
    const text = node => typeof node === 'string' ? node : (node.children ?? []).map(text).join('');
    return button.props.role !== 'tab' && text(button) === label;
  });
  assert.equal(found.length, 1, `one ribbon button labelled ${label}`);
  return found[0];
}

test('PdfEditor mounts, reports busy, navigates pages, and applies an add-page command', async () => {
  assert.equal(fs.existsSync(path.resolve(__dirname, '../src/pdf-editor.css')), true);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/PdfEditor.tsx'), 'utf8');
  assert.equal(/from ['"]\.\/OfficeEditor['"]/.test(source), false);

  let pageCount = 2, opens = 0, renders = 0, renderer;
  const applied = [], changes = [], busy = [];
  const bytes = new Uint8Array([0x25, 0x50, 0x44, 0x46]);
  global.document = { baseURI: 'https://injoffice.invalid/' };
  global.window = { devicePixelRatio: 1 };
  globalThis.__inspectPdf = async () => summary(pageCount);
  globalThis.__applyPdfCommand = async (current, command) => {
    applied.push(command);
    if (command.kind === 'add-page') pageCount += 1;
    return Uint8Array.from([...current, pageCount]);
  };
  globalThis.__pdfView = {
    getDocument() {
      opens += 1;
      const numPages = pageCount;
      return {
        promise: Promise.resolve({
          numPages,
          getPage: async () => ({
            getViewport: () => ({
              width: 400,
              height: 500,
              convertToPdfPoint: (x, y) => [x, y],
              convertToViewportPoint: (x, y) => [x, y],
            }),
            getAnnotations: async () => [],
            getTextContent: async () => ({ items: [{ str: 'Page text' }] }),
            render: () => {
              renders += 1;
              return { promise: Promise.resolve(), cancel() {} };
            },
          }),
        }),
        destroy: async () => {},
      };
    },
  };

  try {
    const Editor = await loadEditor();
    await act(async () => {
      renderer = create(React.createElement(Editor, {
        name: 'Test.pdf',
        bytes,
        onChange: value => changes.push(value),
        onBusyChange: value => busy.push(value),
        viewOptions: { zoom: 100, navigation: true, focus: false },
      }), { createNodeMock: nodeMock });
    });

    const pageNext = () => renderer.root.findByProps({ 'aria-label': 'Next page' });
    await until(() => opens === 1 && !pageNext().props.disabled);
    assert.equal(renderer.root.findByProps({ 'aria-label': 'PDF editor: Test.pdf' }) != null, true);
    assert.equal(busy.includes(true), true);
    assert.equal(busy.at(-1), false);

    await act(async () => pageNext().props.onClick());
    await until(() => renders === 2 && !renderer.root.findByProps({ 'aria-label': 'Previous page' }).props.disabled);
    assert.equal(opens, 1);

    const addPage = ribbonButton(renderer, 'Add page');
    await act(async () => addPage.props.onClick());
    await until(() => opens === 2 && changes.length === 1 && busy.at(-1) === false);
    assert.deepEqual(applied, [{ kind: 'add-page', page: 2 }]);
    assert.equal(pageCount, 3);
    await act(async () => renderer.unmount());
    renderer = undefined;
  } finally {
    if (renderer) await act(async () => renderer.unmount());
    delete globalThis.__pdfView;
    delete globalThis.__inspectPdf;
    delete globalThis.__applyPdfCommand;
    delete global.document;
    delete global.window;
  }
});

test('the PDF editor uses the shared Office ribbon: tabs, icons, find toggle and an export dialog', async () => {
  let renderer;
  const bytes = new Uint8Array([0x25, 0x50, 0x44, 0x46]);
  global.document = { baseURI: 'https://injoffice.invalid/' };
  global.window = { devicePixelRatio: 1, injDesktop: { pickAsset: async () => null, exportBytes: async () => null } };
  globalThis.__inspectPdf = async () => summary(3);
  globalThis.__applyPdfCommand = async current => current;
  globalThis.__pdfView = {
    getDocument: () => ({
      promise: Promise.resolve({
        numPages: 3,
        getPage: async () => ({
          getViewport: () => ({ width: 400, height: 500, convertToPdfPoint: (x, y) => [x, y], convertToViewportPoint: (x, y) => [x, y] }),
          getAnnotations: async () => [],
          getTextContent: async () => ({ items: [{ str: 'Page text' }] }),
          render: () => ({ promise: Promise.resolve(), cancel() {} }),
        }),
      }),
      destroy: async () => {},
    }),
  };
  try {
    const Editor = await loadEditor();
    await act(async () => {
      renderer = create(React.createElement(Editor, { name: 'Test.pdf', bytes, onChange() {}, viewOptions: { zoom: 100, navigation: true, focus: false } }), { createNodeMock: nodeMock });
    });
    await until(() => !renderer.root.findByProps({ 'aria-label': 'Next page' }).props.disabled);

    const label = node => typeof node === 'string' ? node : (node.children ?? []).map(label).join('');
    assert.deepEqual(renderer.root.findAllByProps({ role: 'tab' }).map(label), ['Home', 'Insert', 'View'], 'Office tab strip (File comes from the workspace context)');
    assert.deepEqual(renderer.root.findAllByProps({ role: 'group' }).map(group => group.props['aria-label']),
      ['Tools', 'Editing', 'Text', 'Markup', 'Shapes', 'Illustrations', 'Pages', 'Arrange', 'Page Navigation', 'Export', 'Help']);

    // Every command button carries a glyph and a tooltip; the title-bar Quick Access owns undo/redo.
    const commands = renderer.root.findAllByType('button').filter(button => (button.props.className ?? '').includes('ribbon-button'));
    assert.ok(commands.length >= 15);
    for (const button of commands) {
      assert.equal(button.findAllByType('svg').length, 1, `${label(button) || button.props['aria-label']} has an icon`);
      assert.ok(button.props.title, `${label(button) || button.props['aria-label']} has a tooltip`);
    }
    assert.deepEqual(commands.filter(button => ['Undo', 'Redo'].includes(label(button))), []);
    assert.deepEqual(renderer.root.findAllByProps({ role: 'toolbar' }), [], 'no quick access strip inside the editor');
    assert.equal(renderer.root.findAllByProps({ className: 'pdf-search' }).length, 0, 'the find bar is closed by default');

    // Find opens the search bar with the All pages toggle; closing it puts the chrome back.
    await act(async () => ribbonButton(renderer, 'Find').props.onClick());
    const findBar = renderer.root.findByProps({ 'aria-label': 'Find text in this PDF' });
    assert.equal(findBar.props.className, 'pdf-search');
    const allPages = renderer.root.findAllByType('button').find(button => button.props['aria-label'] === 'All pages' || label(button) === 'All pages');
    assert.equal(allPages.props['aria-pressed'], false);
    await act(async () => allPages.props.onClick());
    assert.equal(renderer.root.findAllByType('button').find(button => label(button) === 'All pages').props['aria-pressed'], true);
    await act(async () => ribbonButton(renderer, 'Find').props.onClick());
    assert.equal(renderer.root.findAllByProps({ className: 'pdf-search' }).length, 0);

    // Export pages is a dialog carrying the range hint, not a permanent row.
    assert.equal(renderer.root.findAllByType('dialog').length, 0);
    await act(async () => ribbonButton(renderer, 'Export pages…').props.onClick());
    const dialog = renderer.root.findByType('dialog');
    assert.equal(dialog.props.className, 'pdf-export-dialog');
    assert.match(label(dialog), /For example: 1-3, 5/);
    assert.equal(renderer.root.findByProps({ 'aria-label': 'PDF pages to export' }).props.placeholder, '1');
    await act(async () => renderer.root.findByType('dialog').props.onCancel({ preventDefault() {} }));

    // One status row: the page on the left, the tool hint on the right, and no Editing support disclosure.
    const status = renderer.root.findByProps({ 'aria-label': 'PDF status' });
    const parts = status.children.map(child => child.children.join(''));
    assert.deepEqual(parts, ['Page 1 of 3', 'Choose a tool to add content or arrange pages.']);
    assert.equal(status.findAllByType('details').length, 0);
    await act(async () => ribbonButton(renderer, 'Editing support').props.onClick());
    assert.match(label(renderer.root.findByType('dialog')), /Editing support/);
    await act(async () => renderer.unmount());
    renderer = undefined;
  } finally {
    if (renderer) await act(async () => renderer.unmount());
    delete globalThis.__pdfView;
    delete globalThis.__inspectPdf;
    delete globalThis.__applyPdfCommand;
    delete global.document;
    delete global.window;
  }
});

test('right-clicking the PDF page opens the shared context menu with the page commands', async () => {
  let renderer;
  const bytes = new Uint8Array([0x25, 0x50, 0x44, 0x46]);
  const listeners = [];
  global.document = { baseURI: 'https://injoffice.invalid/' };
  global.window = {
    devicePixelRatio: 1,
    innerWidth: 1440, innerHeight: 900,
    getSelection: () => null,
    addEventListener() {}, removeEventListener() {},
    document: { addEventListener: (...args) => listeners.push(args[0]), removeEventListener() {}, activeElement: null },
  };
  globalThis.__inspectPdf = async () => summary(1);
  globalThis.__applyPdfCommand = async current => current;
  globalThis.__pdfView = {
    getDocument: () => ({
      promise: Promise.resolve({
        numPages: 1,
        getPage: async () => ({
          getViewport: () => ({ width: 400, height: 500, convertToPdfPoint: (x, y) => [x, 500 - y], convertToViewportPoint: (x, y) => [x, y] }),
          getAnnotations: async () => [],
          getTextContent: async () => ({ items: [{ str: 'Page text' }] }),
          render: () => ({ promise: Promise.resolve(), cancel() {} }),
        }),
      }),
      destroy: async () => {},
    }),
  };
  try {
    const Editor = await loadEditor();
    await act(async () => {
      renderer = create(React.createElement(Editor, { name: 'Test.pdf', bytes, onChange() {}, viewOptions: { zoom: 100, navigation: true, focus: false } }), {
        createNodeMock: element => element.props.className?.startsWith('pdf-page-surface') ? { getBoundingClientRect: () => ({ left: 0, top: 0, width: 400, height: 500 }) } : nodeMock(element),
      });
    });
    await until(() => renderer.root.findAllByProps({ className: 'pdf-text-layer' }).length > 0);
    const surface = renderer.root.findAll(node => typeof node.type === 'string' && (node.props.className ?? '').startsWith('pdf-page-surface'))[0];
    assert.equal(renderer.root.findAllByProps({ role: 'menu' }).length, 0);

    await act(async () => surface.props.onContextMenu({
      preventDefault() {}, clientX: 40, clientY: 60,
      currentTarget: { getBoundingClientRect: () => ({ left: 0, top: 0, width: 400, height: 500 }) },
    }));
    const menu = renderer.root.findByProps({ role: 'menu' });
    assert.equal(menu.props['aria-label'], 'PDF page');
    const items = menu.findAllByType('button').map(button => button.findByProps({ className: 'context-menu-label' }).children.join(''));
    assert.deepEqual(items, ['Select', 'Edit text', 'Add text here', 'Add note', 'Highlight', 'Copy']);
    assert.equal(menu.findAllByProps({ role: 'separator' }).length, 2);
    // Copy needs a page selection; the tools do not.
    const entry = label => menu.findAllByType('button').find(button => button.findByProps({ className: 'context-menu-label' }).children.join('') === label);
    assert.equal(entry('Copy').props['aria-disabled'], true);
    assert.equal(entry('Select').props['aria-checked'], true);
    assert.equal(entry('Add text here').props['aria-disabled'], undefined);

    // "Add text here" picks the tool and places the draft where the menu was opened.
    await act(async () => entry('Add text here').props.onClick());
    assert.equal(renderer.root.findAllByProps({ role: 'menu' }).length, 0, 'the menu closes when an entry runs');
    assert.equal(renderer.root.findByProps({ 'aria-label': 'Tool settings' }) != null, true);
    assert.equal(renderer.root.findByProps({ 'aria-label': 'New PDF text' }) != null, true);
    await act(async () => renderer.unmount());
    renderer = undefined;
  } finally {
    if (renderer) await act(async () => renderer.unmount());
    delete globalThis.__pdfView;
    delete globalThis.__inspectPdf;
    delete globalThis.__applyPdfCommand;
    delete global.document;
    delete global.window;
  }
});
