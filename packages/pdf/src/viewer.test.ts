import { PDFDocument, StandardFonts } from 'pdf-lib';
import { afterEach, describe, expect, it, vi } from 'vitest';
import * as pdfjsLib from 'pdfjs-dist/legacy/build/pdf.mjs';
import { canvasRenderMetrics, configurePdfWorker, PdfViewerDocument, renderPageToCanvas } from './viewer.js';

async function makeDoc(pages: Array<{ width: number; height: number; text: string }>): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  for (const spec of pages) {
    const page = doc.addPage([spec.width, spec.height]);
    page.drawText(spec.text, { x: 20, y: spec.height - 40, size: 18, font });
  }
  return doc.save();
}

describe('PdfViewerDocument', () => {
  let handles: PdfViewerDocument[] = [];

  async function load(bytes: Uint8Array): Promise<PdfViewerDocument> {
    const doc = await PdfViewerDocument.load(bytes);
    handles.push(doc);
    return doc;
  }

  afterEach(async () => {
    await Promise.all(handles.map((h) => h.destroy()));
    handles = [];
  });

  it('reports page count and geometry', async () => {
    const bytes = await makeDoc([
      { width: 210, height: 297, text: 'first' },
      { width: 210, height: 297, text: 'second' },
    ]);
    const doc = await load(bytes);
    expect(doc.pageCount).toBe(2);
    expect(await doc.getPageSize(1)).toEqual({ width: 210, height: 297 });
  });

  it('extracts per-page text', async () => {
    const bytes = await makeDoc([
      { width: 200, height: 200, text: 'hello world' },
      { width: 200, height: 200, text: 'goodbye world' },
    ]);
    const doc = await load(bytes);
    expect(await doc.getPageText(1)).toContain('hello world');
    expect(await doc.getPageText(2)).toContain('goodbye world');
  });

  it('throws for a page index outside the document', async () => {
    const doc = await load(await makeDoc([{ width: 100, height: 100, text: 'only' }]));
    await expect(doc.getPage(0)).rejects.toThrow(/out of range/);
    await expect(doc.getPage(2)).rejects.toThrow(/out of range/);
  });

  it('search finds matches across pages, case-insensitively, with a snippet', async () => {
    const bytes = await makeDoc([
      { width: 300, height: 100, text: 'the Quick Brown Fox' },
      { width: 300, height: 100, text: 'jumps over the lazy dog' },
    ]);
    const doc = await load(bytes);
    const matches = await doc.search('quick');
    expect(matches).toHaveLength(1);
    expect(matches[0].pageIndex).toBe(1);
    expect(matches[0].snippet.toLowerCase()).toContain('quick');
  });

  it('search returns nothing for a blank query', async () => {
    const doc = await load(await makeDoc([{ width: 100, height: 100, text: 'anything' }]));
    expect(await doc.search('   ')).toEqual([]);
  });

  it('reuses extracted text between searches', async () => {
    const doc = await load(await makeDoc([{ width: 200, height: 100, text: 'hello world' }]));
    const page = await doc.getPage(1);
    const extract = vi.spyOn(page, 'getTextContent');
    await doc.search('hello');
    await doc.search('world');
    expect(extract).toHaveBeenCalledTimes(1);
  });

  it('stops a cancelled search before reading the next page', async () => {
    const doc = await load(await makeDoc([{ width: 200, height: 100, text: 'one' }, { width: 200, height: 100, text: 'two' }]));
    const controller = new AbortController();
    const read = vi.spyOn(doc, 'getPageText').mockImplementation(async () => {
      controller.abort();
      return 'one';
    });
    await expect(doc.search('one', { signal: controller.signal })).rejects.toMatchObject({ name: 'AbortError' });
    expect(read).toHaveBeenCalledTimes(1);
    await expect(doc.search('two', { signal: controller.signal })).rejects.toMatchObject({ name: 'AbortError' });
    expect(read).toHaveBeenCalledTimes(1);
  });

  it('rejects fractional and non-finite page indices', async () => {
    const doc = await load(await makeDoc([{ width: 100, height: 100, text: 'one' }]));
    for (const page of [NaN, Infinity, 1.5]) await expect(doc.getPage(page)).rejects.toThrow(/out of range/);
  });

  it('preserves bookmark nesting and resolves zero-based numeric destinations', async () => {
    const doc = await load(await makeDoc([{ width: 100, height: 100, text: 'one' }]));
    const proxy = (doc as unknown as { proxy: pdfjsLib.PDFDocumentProxy }).proxy;
    vi.spyOn(proxy, 'getOutline').mockResolvedValue([
      { title: 'Chapter', dest: [0], items: [{ title: 'Section', dest: [0], items: [] }] },
      { title: 'Invalid target', dest: [99], items: [] },
    ] as unknown as Awaited<ReturnType<typeof proxy.getOutline>>);
    expect(await doc.getOutlineTree()).toEqual([
      { title: 'Chapter', pageIndex: 1, children: [{ title: 'Section', pageIndex: 1, children: [] }] },
      { title: 'Invalid target', pageIndex: null, children: [] },
    ]);
    expect(await doc.getOutline()).toEqual([
      { title: 'Chapter', pageIndex: 1 }, { title: 'Section', pageIndex: 1 }, { title: 'Invalid target', pageIndex: null },
    ]);
  });
});

describe('configurePdfWorker', () => {
  it('sets the worker URL used by pdf.js', () => {
    const previous = pdfjsLib.GlobalWorkerOptions.workerSrc;
    try {
      configurePdfWorker('/injoffice/assets/pdf.worker-demo.mjs');
      expect(pdfjsLib.GlobalWorkerOptions.workerSrc).toBe('/injoffice/assets/pdf.worker-demo.mjs');
    } finally {
      pdfjsLib.GlobalWorkerOptions.workerSrc = previous;
    }
  });

  it('rejects an empty worker URL', () => {
    expect(() => configurePdfWorker('   ')).toThrow(/must not be empty/);
  });
});

describe('canvasRenderMetrics', () => {
  it('keeps CSS size at the requested zoom while scaling the backing store', () => {
    expect(canvasRenderMetrics(612.25, 792.5, 2)).toEqual({
      pixelWidth: 1225,
      pixelHeight: 1585,
      cssWidth: 612.25,
      cssHeight: 792.5,
      transform: [2, 0, 0, 2, 0, 0],
    });
  });

  it('does not add a transform at one device pixel per CSS pixel', () => {
    expect(canvasRenderMetrics(200, 300)).toEqual({
      pixelWidth: 200,
      pixelHeight: 300,
      cssWidth: 200,
      cssHeight: 300,
      transform: undefined,
    });
  });

  it('rejects invalid viewport dimensions and pixel ratios', () => {
    expect(() => canvasRenderMetrics(0, 300, 1)).toThrow(/viewport dimensions/);
    expect(() => canvasRenderMetrics(200, 300, Number.NaN)).toThrow(/pixel ratio/);
  });

  it('caps raster area and dimensions while preserving CSS zoom', () => {
    for (const [width, height, budget] of [[20000, 30000, 16000000], [1000000, 1, 100], [10.1, 10.1, 110], [10000, 1, 1]]) {
      const metrics = canvasRenderMetrics(width, height, 3, budget);
      expect(metrics.pixelWidth * metrics.pixelHeight).toBeLessThanOrEqual(budget);
      expect(Math.max(metrics.pixelWidth, metrics.pixelHeight)).toBeLessThanOrEqual(16384);
      expect(metrics.cssWidth).toBe(width);
      expect(metrics.cssHeight).toBe(height);
    }
    expect(() => canvasRenderMetrics(200, 300, 1, 0)).toThrow(/pixel budget/);
  });
});

describe('render cancellation', () => {
  it('cancels the active PDF.js task and releases the abort listener', async () => {
    const controller = new AbortController();
    let reject!: (reason: Error) => void;
    const task = { promise: new Promise<void>((_resolve, fail) => { reject = fail; }), cancel: vi.fn(() => reject(new Error('cancelled'))) };
    const page = { getViewport: () => ({ width: 100, height: 200 }), render: vi.fn(() => task) };
    const doc = { getPage: async () => page } as unknown as PdfViewerDocument;
    const canvas = { getContext: () => ({}), style: {} } as unknown as HTMLCanvasElement;
    const remove = vi.spyOn(controller.signal, 'removeEventListener');
    const rendering = renderPageToCanvas(doc, 1, canvas, 1, 2, { signal: controller.signal });
    await Promise.resolve();
    controller.abort();
    await expect(rendering).rejects.toThrow('cancelled');
    expect(task.cancel).toHaveBeenCalledTimes(1);
    expect(remove).toHaveBeenCalledWith('abort', expect.any(Function));
    expect(canvas.width).toBe(200);
  });
});
