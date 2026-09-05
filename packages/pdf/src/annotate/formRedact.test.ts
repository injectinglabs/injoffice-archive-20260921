import { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFRawStream, PDFRef, PDFStream } from 'pdf-lib';
import { describe, expect, it } from 'vitest';
import { renderPageToPng } from '../ocr/renderPage.js';
import { PdfViewerDocument } from '../viewer.js';
import { listRedactableFormFields, redactFormFields } from './formRedact.js';

const N = {
  Annots: PDFName.of('Annots'),
  AP: PDFName.of('AP'),
  Kids: PDFName.of('Kids'),
  Subtype: PDFName.of('Subtype'),
  T: PDFName.of('T'),
  Widget: PDFName.of('Widget'),
} as const;

interface Fixture {
  bytes: Uint8Array;
  /** The two widgets deliberately display the same sensitive value on separate pages. */
  secretAppearanceStreams: string[];
}

// pdf-lib may represent /AP /N as either one indirect stream or a state
// dictionary containing one or more streams. Both are valid PDF appearance
// forms, so the fixture must capture the raw streams recursively instead of
// assuming a producer-specific direct /N -> stream shape.
function appearanceStreamSignatures(doc: PDFDocument, object: unknown, seen = new Set<unknown>()): string[] {
  if (!object || seen.has(object)) return [];
  seen.add(object);
  if (object instanceof PDFRef) return appearanceStreamSignatures(doc, doc.context.lookup(object), seen);
  if (object instanceof PDFRawStream) return [Buffer.from(object.getContents()).toString('base64')];
  if (object instanceof PDFDict) {
    const streams: string[] = [];
    for (const [, value] of object.entries()) streams.push(...appearanceStreamSignatures(doc, value, seen));
    return streams;
  }
  if (object instanceof PDFStream) return [];
  return [];
}

async function makeFixture(): Promise<Fixture> {
  const doc = await PDFDocument.create();
  const first = doc.addPage([300, 200]);
  const second = doc.addPage([300, 200]);
  const form = doc.getForm();
  const secret = form.createTextField('taxId');
  secret.setText('REDACTED-FORM-VALUE-742');
  secret.addToPage(first, { x: 20, y: 130, width: 210, height: 25 });
  secret.addToPage(second, { x: 20, y: 90, width: 210, height: 25 });
  const survivor = form.createTextField('publicName');
  survivor.setText('Public name remains');
  survivor.addToPage(first, { x: 20, y: 70, width: 210, height: 25 });
  form.updateFieldAppearances();

  // Capture proof signatures after serialization/reload. At construction
  // time pdf-lib uses PDFContentStream objects, whereas the redaction proof
  // intentionally inspects the raw streams present in actual saved bytes.
  const bytes = await doc.save({ updateFieldAppearances: false });
  const reloaded = await PDFDocument.load(bytes);
  const reloadedSecret = reloaded.getForm().getTextField('taxId');
  const secretAppearanceStreams: string[] = [];
  for (const widget of reloadedSecret.acroField.getWidgets()) {
    const ap = widget.dict.lookupMaybe(N.AP, PDFDict);
    const normal = ap?.get(PDFName.of('N'));
    const streams = appearanceStreamSignatures(reloaded, normal);
    if (streams.length === 0) throw new Error('fixture normal appearance has no raw stream');
    secretAppearanceStreams.push(...streams);
  }
  return { bytes, secretAppearanceStreams };
}

function fieldNames(bytes: Uint8Array): Promise<string[]> {
  return PDFDocument.load(bytes).then((doc) => doc.getForm().getFields().map((field) => field.getName()));
}

async function widgetFieldNames(bytes: Uint8Array, pageNumber: number): Promise<string[]> {
  const doc = await PDFDocument.load(bytes);
  const page = doc.getPages()[pageNumber - 1]!;
  const annots = page.node.lookupMaybe(N.Annots, PDFArray);
  if (!annots) return [];
  const names: string[] = [];
  for (let i = 0; i < annots.size(); i++) {
    const ref = annots.get(i);
    if (!(ref instanceof PDFRef)) continue;
    const dict = doc.context.lookupMaybe(ref, PDFDict);
    if (!dict || dict.lookupMaybe(N.Subtype, PDFName) !== N.Widget) continue;
    const parent = dict.get(PDFName.of('Parent'));
    const parentDict = doc.context.lookupMaybe(parent, PDFDict);
    const t = parentDict?.lookupMaybe(N.T, PDFHexString);
    if (t) names.push(t.decodeText());
  }
  return names;
}

async function allRawStreamSignatures(bytes: Uint8Array): Promise<Set<string>> {
  const doc = await PDFDocument.load(bytes);
  const out = new Set<string>();
  for (const [, object] of doc.context.enumerateIndirectObjects()) {
    if (object instanceof PDFRawStream) out.add(Buffer.from(object.getContents()).toString('base64'));
  }
  return out;
}

describe('redactFormFields', () => {
  it('lists only complete terminal field identities without exposing field values', async () => {
    const fixture = await makeFixture();
    const targets = await listRedactableFormFields(fixture.bytes);
    expect(targets).toEqual([
      { name: 'taxId', wholeField: true },
      { name: 'publicName', wholeField: true },
    ]);
    expect(JSON.stringify(targets)).not.toContain('REDACTED-FORM-VALUE-742');
  });

  it('physically removes a complete multi-widget field, its value, and its appearances', async () => {
    const fixture = await makeFixture();
    const result = await redactFormFields(fixture.bytes, [{ name: 'taxId', wholeField: true }]);
    expect(result.removed).toEqual(['taxId']);
    expect(await fieldNames(result.bytes)).toEqual(['publicName']);
    expect(await widgetFieldNames(result.bytes, 1)).toEqual(['publicName']);
    expect(await widgetFieldNames(result.bytes, 2)).toEqual([]);

    const afterViewer = await PdfViewerDocument.load(result.bytes);
    try {
      expect(await afterViewer.getPageText(1)).not.toContain('REDACTED-FORM-VALUE-742');
      expect(await afterViewer.getPageText(2)).not.toContain('REDACTED-FORM-VALUE-742');
    } finally {
      await afterViewer.destroy();
    }
    // This is an independent pdfium raster reopen, not merely pdf-lib's structural parse.
    const raster = await renderPageToPng(result.bytes, 1, 1);
    expect(raster.png.length).toBeGreaterThan(100);
    const streams = await allRawStreamSignatures(result.bytes);
    for (const signature of fixture.secretAppearanceStreams) expect(streams).not.toContain(signature);
  });

  it('rejects a partial widget-style selection before mutating any output bytes', async () => {
    const fixture = await makeFixture();
    await expect(
      redactFormFields(fixture.bytes, [{ name: 'taxId', wholeField: false } as unknown as { name: string; wholeField: true }]),
    ).rejects.toThrow(/complete field/);
    expect(await fieldNames(fixture.bytes)).toContain('taxId');
  });

  it('prevalidates the full batch so a missing second field cannot produce a partial redaction', async () => {
    const fixture = await makeFixture();
    await expect(
      redactFormFields(fixture.bytes, [
        { name: 'taxId', wholeField: true },
        { name: 'does.not.exist', wholeField: true },
      ]),
    ).rejects.toThrow(/does not exist/);
    expect(await fieldNames(fixture.bytes)).toContain('taxId');
  });

  it('rejects duplicate field selectors as ambiguous instead of deleting once and claiming two redactions', async () => {
    const fixture = await makeFixture();
    await expect(
      redactFormFields(fixture.bytes, [
        { name: 'taxId', wholeField: true },
        { name: 'taxId', wholeField: true },
      ]),
    ).rejects.toThrow(/more than once/);
  });

  it('is a byte-identical no-op for no selections', async () => {
    const fixture = await makeFixture();
    const result = await redactFormFields(fixture.bytes, []);
    expect(result).toEqual({ bytes: fixture.bytes, removed: [] });
  });
});
