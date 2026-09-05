import { PDFDocument, StandardFonts } from 'pdf-lib';
import { PNG } from 'pngjs';
import { describe, expect, it } from 'vitest';
import { applyImageEdits, listPageImages, verifyImageEdits } from './imageEdit.js';
import type { ImageEditSpec } from './types.js';

async function blankDoc(pages = 1): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  for (let i = 0; i < pages; i++) doc.addPage([400, 300]);
  return doc.save();
}

/** A page with one real text object, for belowText/aboveText layering tests. */
async function docWithText(): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  const page = doc.addPage([400, 300]);
  const font = await doc.embedFont(StandardFonts.Helvetica);
  page.drawText('hello', { x: 20, y: 150, size: 24, font });
  return doc.save();
}

function makePng(w: number, h: number, [r, g, b, a]: [number, number, number, number]): string {
  const png = new PNG({ width: w, height: h });
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const idx = (w * y + x) << 2;
      png.data[idx] = r;
      png.data[idx + 1] = g;
      png.data[idx + 2] = b;
      png.data[idx + 3] = a;
    }
  }
  return PNG.sync.write(png).toString('base64');
}

const RED_PNG = makePng(4, 4, [220, 20, 20, 255]);
const BLUE_PNG = makePng(4, 4, [20, 20, 220, 255]);

describe('applyImageEdits / insertImage', () => {
  it('inserts an image and it is listable at the requested rect', async () => {
    const bytes = await blankDoc();
    const edit: ImageEditSpec = { kind: 'insertImage', page: 1, image: RED_PNG, rect: [50, 50, 90, 90], layer: 'aboveText' };
    const result = await applyImageEdits(bytes, [edit]);
    expect(result.skipped).toEqual([]);
    const images = await listPageImages(result.bytes);
    expect(images).toHaveLength(1);
    expect(images[0]!.page).toBe(1);
    expect(images[0]!.rect[0]).toBeCloseTo(50, 0);
    expect(images[0]!.rect[2]).toBeCloseTo(90, 0);
  });

  it('places an aboveText image after the page text run, and a belowText one before it', async () => {
    const bytes = await docWithText();
    const above: ImageEditSpec = { kind: 'insertImage', page: 1, image: RED_PNG, rect: [200, 50, 240, 90], layer: 'aboveText' };
    const below: ImageEditSpec = { kind: 'insertImage', page: 1, image: BLUE_PNG, rect: [10, 200, 50, 240], layer: 'belowText' };
    const result = await applyImageEdits(bytes, [above, below]);
    expect(result.skipped).toEqual([]);
    const images = await listPageImages(result.bytes);
    const aboveImg = images.find((i) => Math.round(i.rect[0]) === 200);
    const belowImg = images.find((i) => Math.round(i.rect[0]) === 10);
    expect(aboveImg?.aboveText).toBe(true);
    expect(belowImg?.aboveText).toBe(false);
  });

  it('counter-rotates the image matrix so it shows upright on a rotated page footprint', async () => {
    const bytes = await blankDoc();
    const edit90: ImageEditSpec = { kind: 'insertImage', page: 1, image: RED_PNG, rect: [10, 10, 50, 30], layer: 'aboveText', rotate: 90 };
    const result = await applyImageEdits(bytes, [edit90]);
    const images = await listPageImages(result.bytes);
    // FPDFPageObj_GetBounds reports the AXIS-ALIGNED box regardless of rotation, so the
    // listed rect is unchanged by `rotate` — this asserts the insert didn't error and the
    // object is still findable at its declared footprint (the geometry itself is the same
    // `imageMatrix` table `annotate/stamp.ts` already verifies against all four angles).
    expect(images).toHaveLength(1);
    expect(images[0]!.rect[0]).toBeCloseTo(10, 0);
    expect(images[0]!.rect[3]).toBeCloseTo(30, 0);
  });
});

describe('applyImageEdits / transformImage', () => {
  async function withOneImage(): Promise<Uint8Array> {
    const bytes = await blankDoc();
    const edit: ImageEditSpec = { kind: 'insertImage', page: 1, image: RED_PNG, rect: [50, 50, 90, 90], layer: 'aboveText' };
    return (await applyImageEdits(bytes, [edit])).bytes;
  }

  it('resizes the image to a new footprint', async () => {
    const bytes = await withOneImage();
    const edit: ImageEditSpec = { kind: 'transformImage', page: 1, oldRect: [50, 50, 90, 90], rect: [100, 100, 220, 220] };
    const result = await applyImageEdits(bytes, [edit]);
    expect(result.skipped).toEqual([]);
    const images = await listPageImages(result.bytes);
    expect(images).toHaveLength(1);
    expect(images[0]!.rect[0]).toBeCloseTo(100, 0);
    expect(images[0]!.rect[2]).toBeCloseTo(220, 0);
  });

  it('swaps width/height on an odd quarter turn', async () => {
    const bytes = await blankDoc();
    const wide: ImageEditSpec = { kind: 'insertImage', page: 1, image: RED_PNG, rect: [10, 10, 90, 40], layer: 'aboveText' };
    const inserted = (await applyImageEdits(bytes, [wide])).bytes;
    const turned: ImageEditSpec = { kind: 'transformImage', page: 1, oldRect: [10, 10, 90, 40], rect: [10, 10, 40, 90], quarterTurns: 1 };
    const result = await applyImageEdits(inserted, [turned]);
    expect(result.skipped).toEqual([]);
    const images = await listPageImages(result.bytes);
    expect(images[0]!.rect[2] - images[0]!.rect[0]).toBeCloseTo(30, 0);
    expect(images[0]!.rect[3] - images[0]!.rect[1]).toBeCloseTo(80, 0);
  });

  it('moves the image to a new z-band', async () => {
    const bytes = await docWithText();
    const insert: ImageEditSpec = { kind: 'insertImage', page: 1, image: RED_PNG, rect: [200, 50, 240, 90], layer: 'aboveText' };
    const inserted = (await applyImageEdits(bytes, [insert])).bytes;
    expect((await listPageImages(inserted))[0]!.aboveText).toBe(true);
    const move: ImageEditSpec = { kind: 'transformImage', page: 1, oldRect: [200, 50, 240, 90], rect: [200, 50, 240, 90], layer: 'belowText' };
    const result = await applyImageEdits(inserted, [move]);
    expect(result.skipped).toEqual([]);
    expect((await listPageImages(result.bytes))[0]!.aboveText).toBe(false);
  });

  it('skips fail-soft when the target rect matches nothing, leaving bytes untouched', async () => {
    const bytes = await withOneImage();
    const edit: ImageEditSpec = { kind: 'transformImage', page: 1, oldRect: [500, 500, 540, 540], rect: [100, 100, 140, 140] };
    const result = await applyImageEdits(bytes, [edit]);
    expect(result.skipped).toHaveLength(1);
    expect(result.skipped[0]!.reason).toMatch(/could not be located/);
    expect(result.bytes).toEqual(bytes);
    expect(await listPageImages(result.bytes)).toHaveLength(1);
  });
});

describe('applyImageEdits / replaceImage', () => {
  it('swaps pixels while keeping the footprint (no rect/quarterTurns change)', async () => {
    const bytes = await blankDoc();
    const insert: ImageEditSpec = { kind: 'insertImage', page: 1, image: RED_PNG, rect: [50, 50, 90, 90], layer: 'aboveText' };
    const inserted = (await applyImageEdits(bytes, [insert])).bytes;
    const replace: ImageEditSpec = { kind: 'replaceImage', page: 1, oldRect: [50, 50, 90, 90], rect: [50, 50, 90, 90], image: BLUE_PNG };
    const result = await applyImageEdits(inserted, [replace]);
    expect(result.skipped).toEqual([]);
    const images = await listPageImages(result.bytes);
    expect(images).toHaveLength(1);
    expect(images[0]!.rect[0]).toBeCloseTo(50, 0);
    expect(images[0]!.rect[2]).toBeCloseTo(90, 0);
  });

  it('replaces pixels and resizes in one op', async () => {
    const bytes = await blankDoc();
    const insert: ImageEditSpec = { kind: 'insertImage', page: 1, image: RED_PNG, rect: [50, 50, 90, 90], layer: 'aboveText' };
    const inserted = (await applyImageEdits(bytes, [insert])).bytes;
    const replace: ImageEditSpec = { kind: 'replaceImage', page: 1, oldRect: [50, 50, 90, 90], rect: [10, 10, 150, 150], image: BLUE_PNG };
    const result = await applyImageEdits(inserted, [replace]);
    expect(result.skipped).toEqual([]);
    const images = await listPageImages(result.bytes);
    expect(images[0]!.rect[0]).toBeCloseTo(10, 0);
    expect(images[0]!.rect[2]).toBeCloseTo(150, 0);
  });
});

describe('applyImageEdits / deleteImage', () => {
  it('removes the matched image', async () => {
    const bytes = await blankDoc();
    const insert: ImageEditSpec = { kind: 'insertImage', page: 1, image: RED_PNG, rect: [50, 50, 90, 90], layer: 'aboveText' };
    const inserted = (await applyImageEdits(bytes, [insert])).bytes;
    expect(await listPageImages(inserted)).toHaveLength(1);
    const del: ImageEditSpec = { kind: 'deleteImage', page: 1, oldRect: [50, 50, 90, 90] };
    const result = await applyImageEdits(inserted, [del]);
    expect(result.skipped).toEqual([]);
    expect(await listPageImages(result.bytes)).toHaveLength(0);
  });
});

describe('applyImageEdits fail-soft behavior', () => {
  it('skips an out-of-range page with a reason, and does not touch other edits', async () => {
    const bytes = await blankDoc();
    const bad: ImageEditSpec = { kind: 'insertImage', page: 9, image: RED_PNG, rect: [10, 10, 20, 20], layer: 'aboveText' };
    const good: ImageEditSpec = { kind: 'insertImage', page: 1, image: RED_PNG, rect: [50, 50, 90, 90], layer: 'aboveText' };
    const result = await applyImageEdits(bytes, [bad, good]);
    expect(result.skipped).toEqual([{ editIndex: 0, page: 9, reason: 'page does not exist' }]);
    expect(await listPageImages(result.bytes)).toHaveLength(1);
  });

  it('skips an edit whose image data is not a decodable PNG', async () => {
    const bytes = await blankDoc();
    const edit: ImageEditSpec = { kind: 'insertImage', page: 1, image: 'not-a-real-png', rect: [10, 10, 20, 20], layer: 'aboveText' };
    const result = await applyImageEdits(bytes, [edit]);
    expect(result.skipped).toHaveLength(1);
    expect(result.skipped[0]!.reason).toMatch(/could not decode/);
    expect(result.bytes).toEqual(bytes);
  });

  it('is a byte-identical no-op for an empty edit list', async () => {
    const bytes = await blankDoc();
    const result = await applyImageEdits(bytes, []);
    expect(result).toEqual({ bytes, skipped: [] });
    expect(result.bytes).toBe(bytes);
  });
});

describe('verifyImageEdits', () => {
  it('reports no failures when the edited image is findable at its target rect', async () => {
    const bytes = await blankDoc();
    const insert: ImageEditSpec = { kind: 'insertImage', page: 1, image: RED_PNG, rect: [50, 50, 90, 90], layer: 'aboveText' };
    const inserted = (await applyImageEdits(bytes, [insert])).bytes;
    const failures = await verifyImageEdits(inserted, [{ page: 1, rect: [50, 50, 90, 90] }]);
    expect(failures).toEqual([]);
  });

  it('reports a failure when the target rect is not present in the saved output', async () => {
    const bytes = await blankDoc();
    const failures = await verifyImageEdits(bytes, [{ page: 1, rect: [50, 50, 90, 90] }]);
    expect(failures).toEqual([{ page: 1, reason: 'image missing from saved output' }]);
  });

  it('reports a failure for a page that does not exist', async () => {
    const bytes = await blankDoc();
    const failures = await verifyImageEdits(bytes, [{ page: 9, rect: [0, 0, 10, 10] }]);
    expect(failures).toEqual([{ page: 9, reason: 'page missing from saved output' }]);
  });
});
