import { PDFDocument, StandardFonts } from 'pdf-lib';
import { describe, expect, it } from 'vitest';
import { applyPageOps, mergeDocuments, normalizeRotation, readInfo, resolveSelector, splitDocument } from './pageOps.js';

async function makeDoc(pages: Array<{ width: number; height: number; label: string }>): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  for (const spec of pages) {
    const page = doc.addPage([spec.width, spec.height]);
    page.drawText(spec.label, { x: 20, y: spec.height - 40, size: 24, font });
  }
  return doc.save();
}

async function labelsOf(bytes: Uint8Array): Promise<string[]> {
  const pdfjsLib = await import('pdfjs-dist/legacy/build/pdf.mjs');
  const task = pdfjsLib.getDocument({ data: bytes });
  const proxy = await task.promise;
  const labels: string[] = [];
  for (let i = 1; i <= proxy.numPages; i++) {
    const page = await proxy.getPage(i);
    const content = await page.getTextContent();
    labels.push(content.items.map((it) => ('str' in it ? it.str : '')).join(''));
  }
  await task.destroy();
  return labels;
}

describe('normalizeRotation', () => {
  it('wraps into 0/90/180/270', () => {
    expect(normalizeRotation(0)).toBe(0);
    expect(normalizeRotation(90)).toBe(90);
    expect(normalizeRotation(360)).toBe(0);
    expect(normalizeRotation(450)).toBe(90);
    expect(normalizeRotation(-90)).toBe(270);
    expect(normalizeRotation(-450)).toBe(270);
  });
});

describe('resolveSelector', () => {
  it("expands 'all' to a 1-indexed run", () => {
    expect(resolveSelector('all', 3)).toEqual([1, 2, 3]);
  });

  it('passes through an explicit list', () => {
    expect(resolveSelector([2, 1], 3)).toEqual([2, 1]);
  });

  it('throws on out-of-range pages', () => {
    expect(() => resolveSelector([0], 3)).toThrow(/out of range/);
    expect(() => resolveSelector([4], 3)).toThrow(/out of range/);
  });
});

describe('readInfo', () => {
  it('reports page count and geometry', async () => {
    const bytes = await makeDoc([
      { width: 200, height: 300, label: 'a' },
      { width: 400, height: 100, label: 'b' },
    ]);
    const info = await readInfo(bytes);
    expect(info.pageCount).toBe(2);
    expect(info.pages).toEqual([
      { index: 1, width: 200, height: 300, rotation: 0 },
      { index: 2, width: 400, height: 100, rotation: 0 },
    ]);
  });
});

describe('applyPageOps: rotate', () => {
  it('adds to the existing rotation and normalizes', async () => {
    const bytes = await makeDoc([{ width: 200, height: 300, label: 'a' }]);
    const once = await applyPageOps(bytes, [{ type: 'rotate', pages: 'all', degrees: 90 }]);
    expect((await readInfo(once)).pages[0].rotation).toBe(90);
    const twice = await applyPageOps(once, [{ type: 'rotate', pages: [1], degrees: 270 }]);
    expect((await readInfo(twice)).pages[0].rotation).toBe(0);
  });
});

describe('applyPageOps: delete', () => {
  it('removes the selected pages and keeps the rest in order', async () => {
    const bytes = await makeDoc([
      { width: 100, height: 100, label: 'a' },
      { width: 100, height: 100, label: 'b' },
      { width: 100, height: 100, label: 'c' },
    ]);
    const out = await applyPageOps(bytes, [{ type: 'delete', pages: [2] }]);
    expect((await readInfo(out)).pageCount).toBe(2);
    expect(await labelsOf(out)).toEqual(['a', 'c']);
  });

  it('refuses to delete every page', async () => {
    const bytes = await makeDoc([{ width: 100, height: 100, label: 'a' }]);
    await expect(applyPageOps(bytes, [{ type: 'delete', pages: [1] }])).rejects.toThrow(/every page/);
  });
});

describe('applyPageOps: reorder', () => {
  it('rebuilds the document in the given order', async () => {
    const bytes = await makeDoc([
      { width: 100, height: 100, label: 'a' },
      { width: 100, height: 100, label: 'b' },
      { width: 100, height: 100, label: 'c' },
    ]);
    const out = await applyPageOps(bytes, [{ type: 'reorder', order: [3, 1, 2] }]);
    expect(await labelsOf(out)).toEqual(['c', 'a', 'b']);
  });

  it('rejects a non-permutation', async () => {
    const bytes = await makeDoc([
      { width: 100, height: 100, label: 'a' },
      { width: 100, height: 100, label: 'b' },
    ]);
    await expect(applyPageOps(bytes, [{ type: 'reorder', order: [1, 1] }])).rejects.toThrow(/permutation/);
  });
});

describe('applyPageOps: insertBlank', () => {
  it('inserts a blank page at the requested position, sized off the neighbor by default', async () => {
    const bytes = await makeDoc([
      { width: 150, height: 200, label: 'a' },
      { width: 150, height: 200, label: 'b' },
    ]);
    const out = await applyPageOps(bytes, [{ type: 'insertBlank', at: 2 }]);
    const info = await readInfo(out);
    expect(info.pageCount).toBe(3);
    expect(info.pages[1]).toMatchObject({ width: 150, height: 200 });
    expect(await labelsOf(out)).toEqual(['a', '', 'b']);
  });

  it('honors an explicit size and rejects an out-of-range position', async () => {
    const bytes = await makeDoc([{ width: 150, height: 200, label: 'a' }]);
    const out = await applyPageOps(bytes, [{ type: 'insertBlank', at: 1, size: { width: 50, height: 60 } }]);
    expect((await readInfo(out)).pages[0]).toMatchObject({ width: 50, height: 60 });
    await expect(applyPageOps(bytes, [{ type: 'insertBlank', at: 5 }])).rejects.toThrow(/out of range/);
  });
});

describe('applyPageOps: crop and resize', () => {
  it('crop shrinks the page to the given box', async () => {
    const bytes = await makeDoc([{ width: 400, height: 400, label: 'a' }]);
    const out = await applyPageOps(bytes, [{ type: 'crop', pages: 'all', box: { x: 0, y: 0, width: 100, height: 150 } }]);
    expect((await readInfo(out)).pages[0]).toMatchObject({ width: 100, height: 150 });
  });

  it('resize stretch hits the exact target size', async () => {
    const bytes = await makeDoc([{ width: 200, height: 100, label: 'a' }]);
    const out = await applyPageOps(bytes, [{ type: 'resize', pages: 'all', width: 300, height: 300, fit: 'stretch' }]);
    expect((await readInfo(out)).pages[0]).toMatchObject({ width: 300, height: 300 });
  });

  it('resize contain hits the exact target size while preserving aspect internally', async () => {
    const bytes = await makeDoc([{ width: 200, height: 100, label: 'a' }]);
    const out = await applyPageOps(bytes, [{ type: 'resize', pages: 'all', width: 300, height: 300, fit: 'contain' }]);
    expect((await readInfo(out)).pages[0]).toMatchObject({ width: 300, height: 300 });
  });
});

describe('applyPageOps: nUp', () => {
  it('packs pages n-to-a-sheet and preserves total content across sheets', async () => {
    const bytes = await makeDoc([
      { width: 200, height: 200, label: 'a' },
      { width: 200, height: 200, label: 'b' },
      { width: 200, height: 200, label: 'c' },
      { width: 200, height: 200, label: 'd' },
    ]);
    const out = await applyPageOps(bytes, [{ type: 'nUp', n: 4 }]);
    expect((await readInfo(out)).pageCount).toBe(1);

    const five = await makeDoc(Array.from({ length: 5 }, (_, i) => ({ width: 100, height: 100, label: String(i) })));
    const outFive = await applyPageOps(five, [{ type: 'nUp', n: 4 }]);
    expect((await readInfo(outFive)).pageCount).toBe(2);
  });
});

describe('a rotate → crop → delete pipeline composes left to right', () => {
  it('applies each op to the previous result', async () => {
    const bytes = await makeDoc([
      { width: 200, height: 200, label: 'a' },
      { width: 200, height: 200, label: 'b' },
    ]);
    const out = await applyPageOps(bytes, [
      { type: 'rotate', pages: [1], degrees: 90 },
      { type: 'crop', pages: [2], box: { x: 0, y: 0, width: 50, height: 50 } },
      { type: 'delete', pages: [1] },
    ]);
    const info = await readInfo(out);
    expect(info.pageCount).toBe(1);
    expect(info.pages[0]).toMatchObject({ width: 50, height: 50 });
  });
});

describe('splitDocument', () => {
  it('produces one output per range', async () => {
    const bytes = await makeDoc([
      { width: 100, height: 100, label: 'a' },
      { width: 100, height: 100, label: 'b' },
      { width: 100, height: 100, label: 'c' },
      { width: 100, height: 100, label: 'd' },
    ]);
    const [first, second] = await splitDocument(bytes, [
      { from: 1, to: 2 },
      { from: 3, to: 4 },
    ]);
    expect(await labelsOf(first)).toEqual(['a', 'b']);
    expect(await labelsOf(second)).toEqual(['c', 'd']);
  });

  it('rejects an inverted or out-of-range range', async () => {
    const bytes = await makeDoc([{ width: 100, height: 100, label: 'a' }]);
    await expect(splitDocument(bytes, [{ from: 2, to: 1 }])).rejects.toThrow(/invalid split range/);
    await expect(splitDocument(bytes, [{ from: 1, to: 5 }])).rejects.toThrow(/invalid split range/);
  });
});

describe('mergeDocuments', () => {
  it('concatenates documents in order', async () => {
    const a = await makeDoc([{ width: 100, height: 100, label: 'a' }]);
    const b = await makeDoc([
      { width: 100, height: 100, label: 'b' },
      { width: 100, height: 100, label: 'c' },
    ]);
    const merged = await mergeDocuments([a, b]);
    expect((await readInfo(merged)).pageCount).toBe(3);
    expect(await labelsOf(merged)).toEqual(['a', 'b', 'c']);
  });

  it('round trips split then merge back to the original page sequence', async () => {
    const bytes = await makeDoc([
      { width: 100, height: 100, label: 'a' },
      { width: 100, height: 100, label: 'b' },
      { width: 100, height: 100, label: 'c' },
    ]);
    const parts = await splitDocument(bytes, [
      { from: 1, to: 1 },
      { from: 2, to: 3 },
    ]);
    const rejoined = await mergeDocuments(parts);
    expect(await labelsOf(rejoined)).toEqual(['a', 'b', 'c']);
  });
});
