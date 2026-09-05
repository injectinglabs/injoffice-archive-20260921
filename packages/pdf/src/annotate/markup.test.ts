import { PDFArray, PDFDict, PDFDocument, PDFName, PDFRawStream, PDFRef, decodePDFRawStream } from 'pdf-lib';
import { describe, expect, it } from 'vitest';
import { applyMarkups } from './markup.js';
import type { MarkupSpec } from './types.js';

async function blankDoc(pages = 1): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  for (let i = 0; i < pages; i++) doc.addPage([400, 300]);
  return doc.save();
}

async function annotsOn(bytes: Uint8Array, pageIndex: number): Promise<{ dict: PDFDict; doc: PDFDocument }[]> {
  const doc = await PDFDocument.load(bytes);
  const page = doc.getPages()[pageIndex - 1]!;
  const arr = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray);
  if (!arr) return [];
  const out: { dict: PDFDict; doc: PDFDocument }[] = [];
  for (let i = 0; i < arr.size(); i++) {
    const ref = arr.get(i);
    if (ref instanceof PDFRef) {
      const dict = doc.context.lookupMaybe(ref, PDFDict);
      if (dict) out.push({ dict, doc });
    }
  }
  return out;
}

function apStreamText(dict: PDFDict, doc: PDFDocument): string {
  const ap = dict.lookupMaybe(PDFName.of('AP'), PDFDict);
  const n = ap?.get(PDFName.of('N'));
  if (!(n instanceof PDFRef)) throw new Error('no /AP /N');
  const stream = doc.context.lookup(n);
  if (!(stream instanceof PDFRawStream)) throw new Error('/AP /N is not a stream');
  return Buffer.from(decodePDFRawStream(stream).decode()).toString('latin1');
}

describe('applyMarkups', () => {
  it('adds a highlight with a real fill-rect appearance stream', async () => {
    const bytes = await blankDoc();
    const spec: MarkupSpec = { page: 1, type: 'highlight', color: [1, 1, 0], quads: [[10, 90, 100, 90, 10, 80, 100, 80]] };
    const result = await applyMarkups(bytes, [spec]);
    const annots = await annotsOn(result, 1);
    expect(annots).toHaveLength(1);
    const { dict, doc } = annots[0]!;
    expect(dict.lookupMaybe(PDFName.of('Subtype'), PDFName)).toBe(PDFName.of('Highlight'));
    const ap = apStreamText(dict, doc);
    expect(ap).toContain('1 1 0 rg');
    expect(ap).toMatch(/re f/);
  });

  it('adds an underline with a real stroked-line appearance stream', async () => {
    const bytes = await blankDoc();
    const spec: MarkupSpec = { page: 1, type: 'underline', color: [1, 0, 0], quads: [[10, 90, 100, 90, 10, 80, 100, 80]] };
    const result = await applyMarkups(bytes, [spec]);
    const { dict, doc } = (await annotsOn(result, 1))[0]!;
    expect(dict.lookupMaybe(PDFName.of('Subtype'), PDFName)).toBe(PDFName.of('Underline'));
    const ap = apStreamText(dict, doc);
    expect(ap).toContain('1 0 0 RG');
    expect(ap).toMatch(/m .* l S/);
  });

  it('adds a strikeout', async () => {
    const bytes = await blankDoc();
    const spec: MarkupSpec = { page: 1, type: 'strikeout', color: [0, 0, 1], quads: [[10, 90, 100, 90, 10, 80, 100, 80]] };
    const result = await applyMarkups(bytes, [spec]);
    const { dict } = (await annotsOn(result, 1))[0]!;
    expect(dict.lookupMaybe(PDFName.of('Subtype'), PDFName)).toBe(PDFName.of('StrikeOut'));
  });

  it('sets Rect from the union of all quad bounds and preserves QuadPoints', async () => {
    const bytes = await blankDoc();
    const spec: MarkupSpec = {
      page: 1,
      type: 'highlight',
      color: [1, 1, 0],
      quads: [
        [10, 90, 100, 90, 10, 80, 100, 80],
        [10, 70, 50, 70, 10, 60, 50, 60],
      ],
    };
    const result = await applyMarkups(bytes, [spec]);
    const { dict } = (await annotsOn(result, 1))[0]!;
    const rect = dict.lookupMaybe(PDFName.of('Rect'), PDFArray)!;
    expect(rect.size()).toBe(4);
    const nums = Array.from({ length: 4 }, (_, i) => rect.lookup(i));
    expect(nums.map((n: any) => n.asNumber())).toEqual([10, 60, 100, 90]);
    const quadPoints = dict.lookupMaybe(PDFName.of('QuadPoints'), PDFArray)!;
    expect(quadPoints.size()).toBe(16);
  });

  it('adds multiple markups across multiple pages', async () => {
    const bytes = await blankDoc(2);
    const result = await applyMarkups(bytes, [
      { page: 1, type: 'highlight', color: [1, 1, 0], quads: [[10, 90, 100, 90, 10, 80, 100, 80]] },
      { page: 2, type: 'underline', color: [1, 0, 0], quads: [[10, 90, 100, 90, 10, 80, 100, 80]] },
    ]);
    expect(await annotsOn(result, 1)).toHaveLength(1);
    expect(await annotsOn(result, 2)).toHaveLength(1);
  });

  it('appends to an existing /Annots array rather than replacing it', async () => {
    const bytes = await blankDoc();
    const once = await applyMarkups(bytes, [{ page: 1, type: 'highlight', color: [1, 1, 0], quads: [[10, 90, 100, 90, 10, 80, 100, 80]] }]);
    const twice = await applyMarkups(once, [{ page: 1, type: 'underline', color: [1, 0, 0], quads: [[10, 70, 100, 70, 10, 60, 100, 60]] }]);
    expect(await annotsOn(twice, 1)).toHaveLength(2);
  });

  it('is a no-op that returns the same bytes for an empty list', async () => {
    const bytes = await blankDoc();
    const result = await applyMarkups(bytes, []);
    expect(result).toBe(bytes);
  });

  it('throws for a page that does not exist', async () => {
    const bytes = await blankDoc();
    await expect(applyMarkups(bytes, [{ page: 5, type: 'highlight', color: [1, 1, 0], quads: [[0, 0, 1, 0, 0, 1, 1, 1]] }])).rejects.toThrow(/does not exist/);
  });
});
