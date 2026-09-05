import { PDFDocument } from 'pdf-lib';
import { describe, expect, it } from 'vitest';
import { PdfViewerDocument } from '../viewer.js';
import { applyTextInserts } from './insert.js';

async function blankDoc(pages = 1): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  for (let i = 0; i < pages; i++) doc.addPage([400, 300]);
  return doc.save();
}

async function pageText(bytes: Uint8Array, pageIndex: number): Promise<string> {
  const doc = await PdfViewerDocument.load(bytes);
  try {
    return await doc.getPageText(pageIndex);
  } finally {
    await doc.destroy();
  }
}

describe('applyTextInserts', () => {
  it('inserts a single-line text object into a blank page', async () => {
    const bytes = await blankDoc();
    const result = await applyTextInserts(bytes, [{ page: 1, origin: [50, 200], text: 'Hello Nick', fontSize: 18, color: [0, 0, 0] }]);
    expect(result.skipped).toEqual([]);
    expect(result.applied).toBe(1);
    expect(await pageText(result.bytes, 1)).toContain('Hello Nick');
  });

  it('stacks multi-line text as separate objects', async () => {
    const bytes = await blankDoc();
    const result = await applyTextInserts(bytes, [{ page: 1, origin: [50, 200], text: 'line one\nline two', fontSize: 18, color: [0, 0, 0] }]);
    expect(result.applied).toBe(1);
    const text = await pageText(result.bytes, 1);
    expect(text).toContain('line one');
    expect(text).toContain('line two');
  });

  it('inserts non-ASCII text via the fallback font cascade', async () => {
    const bytes = await blankDoc();
    const result = await applyTextInserts(bytes, [{ page: 1, origin: [50, 200], text: 'café', fontSize: 18, color: [0, 0, 0] }]);
    expect(result.skipped).toEqual([]);
    expect(await pageText(result.bytes, 1)).toContain('café');
  });

  it('applies multiple inserts across multiple pages independently', async () => {
    const bytes = await blankDoc(2);
    const result = await applyTextInserts(bytes, [
      { page: 1, origin: [50, 200], text: 'on page one', fontSize: 18, color: [0, 0, 0] },
      { page: 2, origin: [50, 200], text: 'on page two', fontSize: 18, color: [0, 0, 0] },
    ]);
    expect(result.applied).toBe(2);
    expect(await pageText(result.bytes, 1)).toContain('on page one');
    expect(await pageText(result.bytes, 2)).toContain('on page two');
  });

  it('skips empty/whitespace-only inserted text without corrupting the file', async () => {
    const bytes = await blankDoc();
    const result = await applyTextInserts(bytes, [{ page: 1, origin: [50, 200], text: '   ', fontSize: 18, color: [0, 0, 0] }]);
    expect(result.applied).toBe(0);
    expect(result.skipped).toHaveLength(1);
    expect(result.skipped[0]?.reason).toMatch(/empty inserted text/);
    expect(result.bytes).toBe(bytes);
  });

  it('skips an insert targeting a page that does not exist', async () => {
    const bytes = await blankDoc();
    const result = await applyTextInserts(bytes, [{ page: 5, origin: [50, 200], text: 'nope', fontSize: 18, color: [0, 0, 0] }]);
    expect(result.applied).toBe(0);
    expect(result.skipped[0]?.reason).toMatch(/does not exist/);
  });

  it('counter-rotates the text axes for a rotated page (upright text at every rotate value)', async () => {
    const bytes = await blankDoc();
    for (const rotate of [0, 90, 180, 270] as const) {
      const result = await applyTextInserts(bytes, [{ page: 1, origin: [50, 200], text: `rot${rotate}`, fontSize: 18, color: [0, 0, 0], rotate }]);
      expect(result.applied).toBe(1);
      expect(await pageText(result.bytes, 1)).toContain(`rot${rotate}`);
    }
  });

  it('inserts using each curated font (P1.5)', async () => {
    const bytes = await blankDoc();
    for (const font of ['arial', 'times', 'courier'] as const) {
      const result = await applyTextInserts(bytes, [{ page: 1, origin: [50, 200], text: `via ${font}`, fontSize: 18, color: [0, 0, 0], font }]);
      expect(result.skipped).toEqual([]);
      expect(result.applied).toBe(1);
      expect(await pageText(result.bytes, 1)).toContain(`via ${font}`);
    }
  });

  it('inserts with a bold+italic curated font variant', async () => {
    const bytes = await blankDoc();
    const result = await applyTextInserts(bytes, [{ page: 1, origin: [50, 200], text: 'bold italic', fontSize: 18, color: [0, 0, 0], font: 'arial', bold: true, italic: true }]);
    expect(result.skipped).toEqual([]);
    expect(result.applied).toBe(1);
    expect(await pageText(result.bytes, 1)).toContain('bold italic');
  });
});
