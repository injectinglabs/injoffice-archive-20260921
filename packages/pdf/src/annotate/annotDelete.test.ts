import { PDFArray, PDFDict, PDFDocument, PDFName, PDFRef } from 'pdf-lib';
import { describe, expect, it } from 'vitest';
import { applyAnnotDeletes } from './annotDelete.js';
import { applyDrawings } from './drawing.js';
import { applyMarkups } from './markup.js';
import type { AnnotDeleteSpec, DrawingSpec, MarkupSpec } from './types.js';

async function blankDoc(pages = 1): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  for (let i = 0; i < pages; i++) doc.addPage([400, 300]);
  return doc.save();
}

interface AnnotRow {
  ref: PDFRef;
  dict: PDFDict;
}

async function annotsOn(bytes: Uint8Array, page = 1): Promise<AnnotRow[]> {
  const doc = await PDFDocument.load(bytes);
  const p = doc.getPages()[page - 1]!;
  const arr = p.node.lookupMaybe(PDFName.of('Annots'), PDFArray);
  if (!arr) return [];
  const out: AnnotRow[] = [];
  for (let i = 0; i < arr.size(); i++) {
    const ref = arr.get(i);
    if (ref instanceof PDFRef) {
      const dict = doc.context.lookupMaybe(ref, PDFDict);
      if (dict) out.push({ ref, dict });
    }
  }
  return out;
}

function contentsOf(dict: PDFDict): string | undefined {
  const c = dict.lookup(PDFName.of('Contents'));
  return c ? (c as { decodeText?(): string }).decodeText?.() : undefined;
}

describe('applyAnnotDeletes', () => {
  it('removes a markup annotation by object number', async () => {
    const bytes = await blankDoc();
    const spec: MarkupSpec = { page: 1, type: 'highlight', color: [1, 1, 0], quads: [[10, 90, 100, 90, 10, 80, 100, 80]] };
    const withMarkup = await applyMarkups(bytes, [spec]);
    const before = await annotsOn(withMarkup);
    expect(before).toHaveLength(1);

    const del: AnnotDeleteSpec = { page: 1, objNum: before[0]!.ref.objectNumber, subtype: 'highlight', rect: [10, 80, 100, 90] };
    const result = await applyAnnotDeletes(withMarkup, [del]);
    expect(await annotsOn(result)).toHaveLength(0);
  });

  it('falls back to a rect/subtype scan when the object number is stale', async () => {
    const bytes = await blankDoc();
    const spec: MarkupSpec = { page: 1, type: 'underline', color: [1, 0, 0], quads: [[10, 90, 100, 90, 10, 80, 100, 80]] };
    const withMarkup = await applyMarkups(bytes, [spec]);
    const before = await annotsOn(withMarkup);

    const del: AnnotDeleteSpec = { page: 1, objNum: before[0]!.ref.objectNumber + 999, subtype: 'underline', rect: [10, 80, 100, 90] };
    const result = await applyAnnotDeletes(withMarkup, [del]);
    expect(await annotsOn(result)).toHaveLength(0);
  });

  it('does not delete when subtype disagrees, even with a matching rect', async () => {
    const bytes = await blankDoc();
    const spec: MarkupSpec = { page: 1, type: 'strikeout', color: [0, 0, 0], quads: [[10, 90, 100, 90, 10, 80, 100, 80]] };
    const withMarkup = await applyMarkups(bytes, [spec]);
    const before = await annotsOn(withMarkup);

    const del: AnnotDeleteSpec = { page: 1, objNum: before[0]!.ref.objectNumber + 999, subtype: 'underline', rect: [10, 80, 100, 90] };
    const result = await applyAnnotDeletes(withMarkup, [del]);
    expect(await annotsOn(result)).toHaveLength(1);
  });

  it('leaves bytes unchanged (byte-identical) when nothing matches', async () => {
    const bytes = await blankDoc();
    const spec: MarkupSpec = { page: 1, type: 'highlight', color: [1, 1, 0], quads: [[10, 90, 100, 90, 10, 80, 100, 80]] };
    const withMarkup = await applyMarkups(bytes, [spec]);
    const del: AnnotDeleteSpec = { page: 1, objNum: 9999, subtype: 'highlight', rect: [500, 500, 600, 600] };
    const result = await applyAnnotDeletes(withMarkup, [del]);
    expect(result).toEqual(withMarkup);
  });

  it('leaves non-markup drawing annotations alone — deletion is only modeled for markup/note subtypes', async () => {
    const bytes = await blankDoc();
    const specs: DrawingSpec[] = [
      { page: 1, kind: 'ink', color: [0, 0, 1], width: 2, paths: [[10, 10, 20, 20, 30, 10]] },
      { page: 1, kind: 'rect', color: [0, 1, 0], width: 1, rect: [50, 50, 90, 90] },
    ];
    const withDrawings = await applyDrawings(bytes, specs);
    const before = await annotsOn(withDrawings);
    expect(before).toHaveLength(2);
    const ink = before.find((r) => r.dict.lookupMaybe(PDFName.of('Subtype'), PDFName) === PDFName.of('Ink'))!;

    // Same objNum/rect as the real ink annotation, but `subtype: 'highlight'` (Ink's
    // real pdfium subtype code isn't one this package models for deletion) — must miss.
    const del: AnnotDeleteSpec = { page: 1, objNum: ink.ref.objectNumber, subtype: 'highlight', rect: [10, 10, 30, 20] };
    const untouched = await applyAnnotDeletes(withDrawings, [del]);
    expect(await annotsOn(untouched)).toHaveLength(2);
  });

  it('deletes the right member of a note reply thread by rect + contents identity', async () => {
    const bytes = await blankDoc();
    const root: DrawingSpec = { page: 1, kind: 'note', color: [1, 1, 0], at: [40, 200], contents: 'root comment', localId: 'root' };
    const reply: DrawingSpec = { page: 1, kind: 'note', color: [1, 1, 0], at: [40, 200], contents: 'reply comment', replyToLocalId: 'root' };
    const withNotes = await applyDrawings(bytes, [root, reply]);
    const before = await annotsOn(withNotes);
    expect(before).toHaveLength(2);
    const rootRow = before.find((r) => contentsOf(r.dict) === 'root comment')!;
    const rootRect = (rootRow.dict.lookupMaybe(PDFName.of('Rect'), PDFArray)!.asArray() as unknown as { asNumber(): number }[]).map((n) => n.asNumber()) as [
      number,
      number,
      number,
      number,
    ];

    const del: AnnotDeleteSpec = { page: 1, objNum: rootRow.ref.objectNumber, subtype: 'note', rect: rootRect, contents: 'root comment' };
    const result = await applyAnnotDeletes(withNotes, [del]);
    const after = await annotsOn(result);
    expect(after).toHaveLength(1);
    expect(contentsOf(after[0]!.dict)).toBe('reply comment');
  });

  it('is a no-op for an empty delete list', async () => {
    const bytes = await blankDoc();
    const result = await applyAnnotDeletes(bytes, []);
    expect(result).toBe(bytes);
  });

  it('targets the right page in a multi-page document', async () => {
    const bytes = await blankDoc(2);
    const spec: MarkupSpec = { page: 2, type: 'highlight', color: [1, 0, 1], quads: [[10, 90, 100, 90, 10, 80, 100, 80]] };
    const withMarkup = await applyMarkups(bytes, [spec]);
    const onPage2 = await annotsOn(withMarkup, 2);
    expect(onPage2).toHaveLength(1);

    const del: AnnotDeleteSpec = { page: 1, objNum: onPage2[0]!.ref.objectNumber, subtype: 'highlight', rect: [10, 80, 100, 90] };
    const untouched = await applyAnnotDeletes(withMarkup, [del]);
    expect(await annotsOn(untouched, 2)).toHaveLength(1);

    const del2: AnnotDeleteSpec = { ...del, page: 2 };
    const result = await applyAnnotDeletes(withMarkup, [del2]);
    expect(await annotsOn(result, 2)).toHaveLength(0);
  });
});
