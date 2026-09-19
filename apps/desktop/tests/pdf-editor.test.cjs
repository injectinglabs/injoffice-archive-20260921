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

    const pageNext = () => renderer.root.findByProps({ 'aria-label': 'Next PDF page' });
    await until(() => opens === 1 && !pageNext().props.disabled);
    assert.equal(renderer.root.findByProps({ 'aria-label': 'PDF editor: Test.pdf' }) != null, true);
    assert.equal(busy.includes(true), true);
    assert.equal(busy.at(-1), false);

    await act(async () => pageNext().props.onClick());
    await until(() => renders === 2 && !renderer.root.findByProps({ 'aria-label': 'Previous PDF page' }).props.disabled);
    assert.equal(opens, 1);

    const addPage = renderer.root.findAllByType('button').find(button => button.children.includes('Add page'));
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
