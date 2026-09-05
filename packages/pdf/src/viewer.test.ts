import { PDFDocument, StandardFonts } from 'pdf-lib';
import { afterEach, describe, expect, it } from 'vitest';
import * as pdfjsLib from 'pdfjs-dist/legacy/build/pdf.mjs';
import { canvasRenderMetrics, configurePdfWorker, PdfViewerDocument } from './viewer.js';

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
});
