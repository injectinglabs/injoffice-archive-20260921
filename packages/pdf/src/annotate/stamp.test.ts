import { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFRawStream, PDFRef, decodePDFRawStream, degrees } from 'pdf-lib';
import { describe, expect, it } from 'vitest';
import { applySignatureStamps, applyStamps, VISUAL_SIGNATURE_CONTENT_PREFIX } from './stamp.js';
import type { SignatureStampSpec, StampSpec } from './types.js';

// A real, minimal 1x1 PNG (public-domain test fixture) — exercises pdf-lib's actual
// PNG embedding rather than a mock.
const PNG_1X1 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';

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

describe('applyStamps (burned-in image)', () => {
  it('draws the image into the content stream, not as an annotation', async () => {
    const bytes = await blankDoc();
    const spec: StampSpec = { page: 1, image: PNG_1X1, rect: [10, 10, 100, 60] };
    const result = await applyStamps(bytes, [spec]);
    // No annotation was created for a burned-in stamp.
    expect(await annotsOn(result, 1)).toEqual([]);
    const loaded = await PDFDocument.load(result);
    expect(loaded.getPages()[0]!.node.Resources()?.lookupMaybe(PDFName.of('XObject'), PDFDict)).toBeDefined();
  });

  it('is a no-op that returns the same bytes for an empty list', async () => {
    const bytes = await blankDoc();
    expect(await applyStamps(bytes, [])).toBe(bytes);
  });

  it('throws for a page that does not exist', async () => {
    const bytes = await blankDoc();
    await expect(applyStamps(bytes, [{ page: 5, image: PNG_1X1, rect: [0, 0, 10, 10] }])).rejects.toThrow(/does not exist/);
  });
});

describe('applySignatureStamps (Stamp annotation)', () => {
  it('adds a Stamp annotation with an image-drawing appearance stream', async () => {
    const bytes = await blankDoc();
    const spec: SignatureStampSpec = { page: 1, image: PNG_1X1, rect: [10, 10, 100, 60] };
    const result = await applySignatureStamps(bytes, [spec]);
    const { dict, doc } = (await annotsOn(result, 1))[0]!;
    expect(dict.lookupMaybe(PDFName.of('Subtype'), PDFName)).toBe(PDFName.of('Stamp'));
    const ap = apStreamText(dict, doc);
    expect(ap).toContain('/Im0 Do');
  });

  it('counter-rotates the image matrix for each of the four page rotations without erroring', async () => {
    for (const rotate of [0, 90, 180, 270] as const) {
      const doc = await PDFDocument.create();
      const page = doc.addPage([400, 300]);
      page.setRotation(degrees(rotate));
      const bytes = await doc.save();
      const result = await applySignatureStamps(bytes, [{ page: 1, image: PNG_1X1, rect: [10, 10, 100, 60] }]);
      const { dict, doc: loadedDoc } = (await annotsOn(result, 1))[0]!;
      expect(apStreamText(dict, loadedDoc)).toContain('cm');
    }
  });

  it('sets visual-signature metadata when formFieldName is given', async () => {
    const bytes = await blankDoc();
    const result = await applySignatureStamps(bytes, [{ page: 1, image: PNG_1X1, rect: [10, 10, 100, 60], formFieldName: 'sig1' }]);
    const { dict } = (await annotsOn(result, 1))[0]!;
    const contents = dict.lookup(PDFName.of('Contents'));
    expect(contents).toBeInstanceOf(PDFHexString);
    expect((contents as PDFHexString).decodeText()).toBe(`${VISUAL_SIGNATURE_CONTENT_PREFIX}sig1`);
    const field = dict.lookup(PDFName.of('InjOfficeFormField'));
    expect((field as PDFHexString).decodeText()).toBe('sig1');
  });

  it('omits visual-signature metadata when formFieldName is absent', async () => {
    const bytes = await blankDoc();
    const result = await applySignatureStamps(bytes, [{ page: 1, image: PNG_1X1, rect: [10, 10, 100, 60] }]);
    const { dict } = (await annotsOn(result, 1))[0]!;
    expect(dict.get(PDFName.of('InjOfficeFormField'))).toBeUndefined();
  });

  it('is a no-op that returns the same bytes for an empty list', async () => {
    const bytes = await blankDoc();
    expect(await applySignatureStamps(bytes, [])).toBe(bytes);
  });

  it('throws for a page that does not exist', async () => {
    const bytes = await blankDoc();
    await expect(applySignatureStamps(bytes, [{ page: 5, image: PNG_1X1, rect: [0, 0, 10, 10] }])).rejects.toThrow(/does not exist/);
  });
});
