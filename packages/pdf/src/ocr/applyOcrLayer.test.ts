import { PDFArray, PDFDocument, PDFName, PDFRawStream, PDFRef, StandardFonts, decodePDFRawStream } from 'pdf-lib';
import { describe, expect, it, vi } from 'vitest';
import { PdfViewerDocument } from '../viewer.js';
import { applyOcrLayer } from './applyOcrLayer.js';
import type { OcrProvider, OcrWord } from './types.js';

async function blankDoc(pages = 1): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  for (let i = 0; i < pages; i++) doc.addPage([300, 200]);
  return doc.save();
}

async function docWithText(): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  const page = doc.addPage([300, 200]);
  const font = await doc.embedFont(StandardFonts.Helvetica);
  page.drawText('already has real text', { x: 20, y: 100, size: 14, font });
  return doc.save();
}

async function pageContentText(bytes: Uint8Array, pageIndex: number): Promise<string> {
  const doc = await PDFDocument.load(bytes);
  const page = doc.getPages()[pageIndex - 1]!;
  const contents = page.node.get(PDFName.of('Contents'));
  const streams: PDFRawStream[] = [];
  const resolved = contents instanceof PDFRef ? doc.context.lookup(contents) : contents;
  if (resolved instanceof PDFRawStream) streams.push(resolved);
  else if (resolved instanceof PDFArray) {
    for (let i = 0; i < resolved.size(); i++) {
      const s = doc.context.lookup(resolved.get(i));
      if (s instanceof PDFRawStream) streams.push(s);
    }
  }
  return streams.map((s) => Buffer.from(decodePDFRawStream(s).decode()).toString('latin1')).join('\n');
}

const wordsProvider = (words: OcrWord[]): OcrProvider => async () => words;

describe('applyOcrLayer', () => {
  it('splices recognized words into a page with no text as real, extractable text', async () => {
    const bytes = await blankDoc();
    const words: OcrWord[] = [
      { text: 'Hello', box: [0.05, 0.1, 0.3, 0.25] },
      { text: 'World', box: [0.35, 0.1, 0.6, 0.25] },
    ];
    const result = await applyOcrLayer(bytes, wordsProvider(words));
    expect(result.pagesOcred).toEqual([1]);
    expect(result.pagesSkippedExistingText).toEqual([]);
    expect(result.failed).toEqual([]);

    const viewer = await PdfViewerDocument.load(result.bytes);
    const text = await viewer.getPageText(1);
    await viewer.destroy();
    expect(text).toContain('Hello');
    expect(text).toContain('World');
  });

  it('writes the text in invisible render mode (real "3 Tr" in the content stream)', async () => {
    const bytes = await blankDoc();
    const words: OcrWord[] = [{ text: 'Ghost', box: [0.1, 0.1, 0.4, 0.3] }];
    const result = await applyOcrLayer(bytes, wordsProvider(words));
    const content = await pageContentText(result.bytes, 1);
    expect(content).toMatch(/\b3 Tr\b/);
  });

  it('never calls the provider for a page that already has real text, and leaves it byte-identical', async () => {
    const bytes = await docWithText();
    const provider = vi.fn(async () => []);
    const result = await applyOcrLayer(bytes, provider);
    expect(provider).not.toHaveBeenCalled();
    expect(result.pagesSkippedExistingText).toEqual([1]);
    expect(result.pagesOcred).toEqual([]);
    expect(result.bytes).toBe(bytes);
  });

  it('is fail-soft per page: one page erroring does not block another', async () => {
    const bytes = await blankDoc(2);
    const provider: OcrProvider = async (input) => {
      if (input.page === 1) throw new Error('vision call failed');
      return [{ text: 'Second', box: [0.1, 0.1, 0.4, 0.3] }];
    };
    const result = await applyOcrLayer(bytes, provider);
    expect(result.failed).toEqual([{ page: 1, reason: 'vision call failed' }]);
    expect(result.pagesOcred).toEqual([2]);
    const viewer = await PdfViewerDocument.load(result.bytes);
    const page2Text = await viewer.getPageText(2);
    await viewer.destroy();
    expect(page2Text).toContain('Second');
  });

  it('reports a page with zero placeable words as failed, not ocred', async () => {
    const bytes = await blankDoc();
    const result = await applyOcrLayer(bytes, wordsProvider([]));
    expect(result.pagesOcred).toEqual([]);
    expect(result.failed).toEqual([{ page: 1, reason: 'no placeable words returned' }]);
    expect(result.bytes).toBe(bytes);
  });

  it('skips a blank-text or degenerate-box word without failing the whole page', async () => {
    const bytes = await blankDoc();
    const words: OcrWord[] = [
      { text: '   ', box: [0.1, 0.1, 0.4, 0.3] },
      { text: 'Real', box: [0.5, 0.5, 0.4, 0.6] }, // degenerate: x2 < x1
      { text: 'Kept', box: [0.1, 0.1, 0.4, 0.3] },
    ];
    const result = await applyOcrLayer(bytes, wordsProvider(words));
    expect(result.pagesOcred).toEqual([1]);
    const viewer = await PdfViewerDocument.load(result.bytes);
    const text = await viewer.getPageText(1);
    await viewer.destroy();
    expect(text).toContain('Kept');
    expect(text).not.toContain('Real');
  });

  it('passes the rendered page dimensions to the provider', async () => {
    const bytes = await blankDoc();
    let seen: { width: number; height: number } | null = null;
    const provider: OcrProvider = async (input) => {
      seen = { width: input.width, height: input.height };
      return [];
    };
    await applyOcrLayer(bytes, provider, 2);
    expect(seen).toEqual({ width: 600, height: 400 });
  });
});
