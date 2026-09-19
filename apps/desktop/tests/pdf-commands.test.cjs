const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

function isExternal(id) {
  return id === 'pdf-lib' || id === '@pdf-lib/fontkit' || id.includes('packages/pdf/src/annotate/') || id.includes('pdfjs-dist');
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
  if (id.includes('packages/pdf/src/annotate/drawing')) return { applyDrawings: async () => new Uint8Array() };
  if (id.includes('packages/pdf/src/annotate/markup')) return { applyMarkups: async () => new Uint8Array() };
  if (id.includes('packages/pdf/src/annotate/pdfObjects')) {
    return { annotationEntries: () => [], nameValue() {}, numberArray() {}, textValue() {} };
  }
  if (id.includes('pdfjs-dist')) return { getDocument() { return { promise: Promise.resolve({ numPages: 0 }) }; } };
  return require(id);
}

async function loadCommands() {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.resolve(__dirname, '../src/pdf-commands.ts'),
    platform: 'node',
    external: isExternal,
  });
  try {
    const { output } = await bundle.generate({ format: 'cjs', codeSplitting: false });
    const mod = { exports: {} };
    new Function('require', 'module', 'exports', output[0].code)(stubRequire, mod, mod.exports);
    return mod.exports;
  } finally {
    await bundle.close();
  }
}

test('parsePdfPageRange selects unique pages and rejects invalid ranges', async () => {
  const { parsePdfPageRange } = await loadCommands();
  assert.deepEqual(parsePdfPageRange('1-2, 2, 3', 3), [1, 2, 3]);
  assert.deepEqual(parsePdfPageRange('1–3, 5', 10), [1, 2, 3, 5]);
  assert.deepEqual(parsePdfPageRange('3, 1', 3), [1, 3]);
  assert.throws(() => parsePdfPageRange('4-6', 3), /between 1 and 3/);
  assert.throws(() => parsePdfPageRange('1-0', 3), /between 1 and 3/);
  assert.throws(() => parsePdfPageRange('pages', 3), /1-3, 5/);
  assert.throws(() => parsePdfPageRange('1', 0), /1–3, 5/);
});

test('PdfHistory undo/redo restores exact bytes and push clears redo', async () => {
  const { PdfHistory } = await loadCommands();
  const original = Uint8Array.from([1, 2, 3]);
  const history = new PdfHistory(original);
  original[0] = 9;
  assert.deepEqual(history.bytes, Uint8Array.from([1, 2, 3]));
  const changed = Uint8Array.from([4, 5]);
  history.push(changed);
  changed[0] = 8;
  assert.deepEqual(history.undo(), Uint8Array.from([1, 2, 3]));
  assert.equal(history.canRedo, true);
  assert.deepEqual(history.redo(), Uint8Array.from([4, 5]));
  history.undo();
  history.push(Uint8Array.from([7]));
  assert.equal(history.canRedo, false);
  assert.deepEqual(history.bytes, Uint8Array.from([7]));
  assert.deepEqual(history.undo(), Uint8Array.from([1, 2, 3]));
  const capped = new PdfHistory(Uint8Array.from([0]));
  for (let i = 1; i <= 21; i++) capped.push(Uint8Array.from([i]));
  while (capped.canUndo) capped.undo();
  assert.deepEqual(capped.bytes, Uint8Array.from([1]));
});

test('findPdfTextMatches locates phrases across parts without changing text', async () => {
  const { findPdfTextMatches } = await loadCommands();
  const parts = ['Hello', 'local world', 'Hello'];
  const matches = findPdfTextMatches(parts, 'hello local');
  assert.equal(matches.length, 1);
  assert.deepEqual(matches[0].spans, [0, 1]);
  assert.equal(findPdfTextMatches(parts, 'hello').length, 2);
  assert.equal(findPdfTextMatches(parts, '   ').length, 0);
  assert.equal(findPdfTextMatches(parts, 'x'.repeat(501)).length, 0);
  assert.deepEqual(parts, ['Hello', 'local world', 'Hello']);
});
