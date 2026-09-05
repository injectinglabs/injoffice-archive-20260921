import { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFRawStream, PDFRef, PDFString, decodePDFRawStream } from 'pdf-lib';
import { describe, expect, it } from 'vitest';
import { applyDrawings, applyNoteEdits } from './drawing.js';
import type { DrawingSpec, NoteEditSpec } from './types.js';

async function blankDoc(pages = 1): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  for (let i = 0; i < pages; i++) doc.addPage([400, 300]);
  return doc.save();
}

async function annotsOn(bytes: Uint8Array, pageIndex: number): Promise<{ dict: PDFDict; ref: PDFRef; doc: PDFDocument }[]> {
  const doc = await PDFDocument.load(bytes);
  const page = doc.getPages()[pageIndex - 1]!;
  const arr = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray);
  if (!arr) return [];
  const out: { dict: PDFDict; ref: PDFRef; doc: PDFDocument }[] = [];
  for (let i = 0; i < arr.size(); i++) {
    const ref = arr.get(i);
    if (ref instanceof PDFRef) {
      const dict = doc.context.lookupMaybe(ref, PDFDict);
      if (dict) out.push({ dict, ref, doc });
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

const contentsOf = (dict: PDFDict): string => {
  const c = dict.lookup(PDFName.of('Contents'));
  return c instanceof PDFString || c instanceof PDFHexString ? c.decodeText() : '';
};

describe('applyDrawings: shapes', () => {
  it('draws an ink stroke with a real path appearance stream and InkList', async () => {
    const bytes = await blankDoc();
    const spec: DrawingSpec = { kind: 'ink', page: 1, color: [0, 0, 0], width: 2, paths: [[10, 10, 20, 20, 30, 10]] };
    const result = await applyDrawings(bytes, [spec]);
    const { dict, doc } = (await annotsOn(result, 1))[0]!;
    expect(dict.lookupMaybe(PDFName.of('Subtype'), PDFName)).toBe(PDFName.of('Ink'));
    expect(apStreamText(dict, doc)).toMatch(/m\n.*l\n.*l\nS/s);
    expect(dict.lookupMaybe(PDFName.of('InkList'), PDFArray)).toBeDefined();
  });

  it('draws a rect (Square annotation)', async () => {
    const bytes = await blankDoc();
    const spec: DrawingSpec = { kind: 'rect', page: 1, color: [1, 0, 0], width: 1, rect: [10, 10, 50, 40] };
    const result = await applyDrawings(bytes, [spec]);
    const { dict, doc } = (await annotsOn(result, 1))[0]!;
    expect(dict.lookupMaybe(PDFName.of('Subtype'), PDFName)).toBe(PDFName.of('Square'));
    expect(apStreamText(dict, doc)).toContain('re S');
  });

  it('draws an ellipse (Circle annotation) via Bezier approximation', async () => {
    const bytes = await blankDoc();
    const spec: DrawingSpec = { kind: 'ellipse', page: 1, color: [0, 1, 0], width: 1, rect: [10, 10, 50, 40] };
    const result = await applyDrawings(bytes, [spec]);
    const { dict, doc } = (await annotsOn(result, 1))[0]!;
    expect(dict.lookupMaybe(PDFName.of('Subtype'), PDFName)).toBe(PDFName.of('Circle'));
    const ap = apStreamText(dict, doc);
    expect(ap).toMatch(/ m/);
    expect((ap.match(/c$/gm) ?? []).length + (ap.match(/ c /g) ?? []).length).toBeGreaterThan(0);
  });

  it('draws a line with the /L endpoints', async () => {
    const bytes = await blankDoc();
    const spec: DrawingSpec = { kind: 'line', page: 1, color: [0, 0, 0], width: 1, from: [10, 10], to: [50, 50] };
    const result = await applyDrawings(bytes, [spec]);
    const { dict } = (await annotsOn(result, 1))[0]!;
    expect(dict.lookupMaybe(PDFName.of('Subtype'), PDFName)).toBe(PDFName.of('Line'));
    const l = dict.lookupMaybe(PDFName.of('L'), PDFArray)!;
    expect(l.size()).toBe(4);
  });

  it('draws an arrow with extra arrowhead strokes in the appearance', async () => {
    const bytes = await blankDoc();
    const spec: DrawingSpec = { kind: 'arrow', page: 1, color: [0, 0, 0], width: 1, from: [10, 10], to: [50, 50] };
    const result = await applyDrawings(bytes, [spec]);
    const { dict, doc } = (await annotsOn(result, 1))[0]!;
    const ap = apStreamText(dict, doc);
    // Main shaft + two arrowhead strokes = 3 "m ... l S" segments.
    expect((ap.match(/ m /g) ?? []).length).toBe(3);
  });
});

describe('applyDrawings: sticky notes and reply threads', () => {
  it('adds a root note as a Text annotation with contents/author/date', async () => {
    const bytes = await blankDoc();
    const spec: DrawingSpec = { kind: 'note', page: 1, color: [1, 1, 0], at: [100, 200], contents: 'Hello', author: 'Nick' };
    const result = await applyDrawings(bytes, [spec]);
    const { dict } = (await annotsOn(result, 1))[0]!;
    expect(dict.lookupMaybe(PDFName.of('Subtype'), PDFName)).toBe(PDFName.of('Text'));
    expect(contentsOf(dict)).toBe('Hello');
    expect(dict.get(PDFName.of('IRT'))).toBeUndefined();
  });

  it('threads a reply onto a note in the SAME batch via localId', async () => {
    const bytes = await blankDoc();
    const result = await applyDrawings(bytes, [
      { kind: 'note', page: 1, color: [1, 1, 0], at: [100, 200], contents: 'root', localId: 'a' },
      { kind: 'note', page: 1, color: [1, 1, 0], at: [100, 150], contents: 'reply', replyToLocalId: 'a' },
    ]);
    const annots = await annotsOn(result, 1);
    expect(annots).toHaveLength(2);
    const [root, reply] = annots;
    const irt = reply!.dict.get(PDFName.of('IRT'));
    expect(irt).toBeInstanceOf(PDFRef);
    expect((irt as PDFRef).objectNumber).toBe(root!.ref.objectNumber);
    expect(reply!.dict.lookupMaybe(PDFName.of('RT'), PDFName)).toBe(PDFName.of('R'));
  });

  it('threads a reply onto a note already SAVED in the file (replyToSaved)', async () => {
    const bytes = await blankDoc();
    const withRoot = await applyDrawings(bytes, [{ kind: 'note', page: 1, color: [1, 1, 0], at: [100, 200], contents: 'root' }]);
    const rootDict = (await annotsOn(withRoot, 1))[0]!.dict;
    const rootRect = rootDict.lookupMaybe(PDFName.of('Rect'), PDFArray)!;
    const rect: [number, number, number, number] = [0, 1, 2, 3].map((i) => (rootRect.lookup(i) as any).asNumber()) as any;

    const result = await applyDrawings(withRoot, [{ kind: 'note', page: 1, color: [1, 1, 0], at: [100, 150], contents: 'reply', replyToSaved: { objNum: 0, rect, contents: 'root' } }]);
    const annots = await annotsOn(result, 1);
    expect(annots).toHaveLength(2);
    const reply = annots.find((a) => contentsOf(a.dict) === 'reply')!;
    expect(reply.dict.get(PDFName.of('IRT'))).toBeDefined();
  });

  it('degrades to a root note (does not throw) when a reply target cannot be resolved', async () => {
    const bytes = await blankDoc();
    const result = await applyDrawings(bytes, [{ kind: 'note', page: 1, color: [1, 1, 0], at: [100, 150], contents: 'orphan reply', replyToSaved: { objNum: 999, rect: [0, 0, 1, 1], contents: 'nonexistent' } }]);
    const { dict } = (await annotsOn(result, 1))[0]!;
    expect(contentsOf(dict)).toBe('orphan reply');
    expect(dict.get(PDFName.of('IRT'))).toBeUndefined();
  });
});

describe('applyDrawings: page validation and no-op', () => {
  it('is a no-op that returns the same bytes for an empty list', async () => {
    const bytes = await blankDoc();
    expect(await applyDrawings(bytes, [])).toBe(bytes);
  });

  it('throws for a page that does not exist', async () => {
    const bytes = await blankDoc();
    await expect(applyDrawings(bytes, [{ kind: 'rect', page: 5, color: [0, 0, 0], width: 1, rect: [0, 0, 1, 1] }])).rejects.toThrow(/does not exist/);
  });
});

describe('applyNoteEdits', () => {
  it('rewrites a saved note\'s contents in place, keeping its object number', async () => {
    const bytes = await blankDoc();
    const withNote = await applyDrawings(bytes, [{ kind: 'note', page: 1, color: [1, 1, 0], at: [100, 200], contents: 'original' }]);
    const before = (await annotsOn(withNote, 1))[0]!;
    const rect = before.dict.lookupMaybe(PDFName.of('Rect'), PDFArray)!;
    const rectVals: [number, number, number, number] = [0, 1, 2, 3].map((i) => (rect.lookup(i) as any).asNumber()) as any;

    const edit: NoteEditSpec = { page: 1, objNum: before.ref.objectNumber, rect: rectVals, oldContents: 'original', contents: 'edited' };
    const result = await applyNoteEdits(withNote, [edit]);
    const after = (await annotsOn(result, 1))[0]!;
    expect(contentsOf(after.dict)).toBe('edited');
    expect(after.ref.objectNumber).toBe(before.ref.objectNumber);
  });

  it('is a silent no-op (not an error) when the identity no longer matches', async () => {
    const bytes = await blankDoc();
    const withNote = await applyDrawings(bytes, [{ kind: 'note', page: 1, color: [1, 1, 0], at: [100, 200], contents: 'original' }]);
    const edit: NoteEditSpec = { page: 1, objNum: 9999, rect: [0, 0, 1, 1], oldContents: 'wrong contents', contents: 'edited' };
    const result = await applyNoteEdits(withNote, [edit]);
    expect(contentsOf((await annotsOn(result, 1))[0]!.dict)).toBe('original');
  });

  it('is a no-op that returns the same bytes for an empty list', async () => {
    const bytes = await blankDoc();
    expect(await applyNoteEdits(bytes, [])).toBe(bytes);
  });
});
