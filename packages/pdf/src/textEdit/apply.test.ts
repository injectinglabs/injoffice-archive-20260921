import { PDFArray, PDFDocument, StandardFonts, decodePDFRawStream, degrees } from 'pdf-lib';
import { describe, expect, it } from 'vitest';
import { readInfo } from '../pageOps.js';
import { PdfViewerDocument } from '../viewer.js';
import { applyTextEdits, editBlock, editText, type BlockEditSpec, type SurgicalEditSpec } from './apply.js';

async function onePageDoc(lines: string[]): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const page = doc.addPage([400, 300]);
  lines.forEach((line, i) => {
    page.drawText(line, { x: 40, y: 250 - i * 40, size: 24, font });
  });
  return doc.save();
}

async function rectFor(bytes: Uint8Array, pageIndex: number, text: string): Promise<[number, number, number, number]> {
  // Approximate the run's rect from pdfjs layout — good enough to clear the
  // ≥50%-overlap bar against pdfium's own (slightly different) ink bounds.
  const doc = await PdfViewerDocument.load(bytes);
  try {
    const page = await doc.getPage(pageIndex);
    const content = await page.getTextContent();
    const item = content.items.find((it) => 'str' in it && it.str === text) as { transform: number[]; width: number; height: number } | undefined;
    if (!item) throw new Error(`text ${JSON.stringify(text)} not found on page ${pageIndex}`);
    const [, , , , x, y] = item.transform;
    return [x, y, x + item.width, y + item.height];
  } finally {
    await doc.destroy();
  }
}

async function pageText(bytes: Uint8Array, pageIndex: number): Promise<string> {
  const doc = await PdfViewerDocument.load(bytes);
  try {
    return await doc.getPageText(pageIndex);
  } finally {
    await doc.destroy();
  }
}

/** Raw decoded content-stream text of a page (1-indexed) — real verification that a
 * style actually landed (e.g. distinct `rg` fill-color operators), not just that the
 * text reads back correctly. */
async function pageContentStreamText(bytes: Uint8Array, pageIndex: number): Promise<string> {
  const doc = await PDFDocument.load(bytes);
  const page = doc.getPages()[pageIndex - 1]!;
  const contents = page.node.Contents();
  const streams = contents instanceof PDFArray ? Array.from({ length: contents.size() }, (_, i) => doc.context.lookup(contents.get(i))) : [contents];
  return streams
    .map((s) => Buffer.from(decodePDFRawStream(s as Parameters<typeof decodePDFRawStream>[0]).decode()).toString('latin1'))
    .join('\n');
}

/** Distinct `r g b rg` (non-stroking fill color) operators used in a content stream. */
function distinctFillColors(streamText: string): Set<string> {
  const colors = new Set<string>();
  for (const m of streamText.matchAll(/(-?[\d.]+)\s+(-?[\d.]+)\s+(-?[\d.]+)\s+rg\b/g)) colors.add(`${m[1]},${m[2]},${m[3]}`);
  return colors;
}

/** Hex-string `Tj` show-text payloads (`<48656C6C6F> Tj`), uppercased — the P1.7
 * keep-plan's whole point is that a KEPT object's glyph encoding survives byte-
 * identical (only its position moves), which a redrawn object (re-subsetted, often
 * different glyph ids even for visually identical text) would not reproduce. */
function hexTjPayloads(streamText: string): string[] {
  return [...streamText.matchAll(/<([0-9A-Fa-f]+)>\s*Tj/g)].map((m) => m[1]!.toUpperCase());
}

describe('applyTextEdits: reuse-font in-place edit', () => {
  it('replaces ASCII text already in the run\'s font subset', async () => {
    const bytes = await onePageDoc(['Hello World']);
    const rect = await rectFor(bytes, 1, 'Hello World');
    const result = await applyTextEdits(bytes, [{ page: 1, rect, oldText: 'Hello World', newText: 'Hello World' }]);
    expect(result.skipped).toEqual([]);
    expect(result.applied).toBe(1);
    expect(await pageText(result.bytes, 1)).toBe('Hello World');
  });

  it('replaces with a shorter string built entirely from characters already drawn', async () => {
    // 'World' is a subset of the chars in 'Hello World' (o, r, l, d already drawn).
    const bytes = await onePageDoc(['Hello World']);
    const rect = await rectFor(bytes, 1, 'Hello World');
    const result = await applyTextEdits(bytes, [{ page: 1, rect, oldText: 'Hello World', newText: 'World' }]);
    expect(result.applied).toBe(1);
    expect(result.skipped).toEqual([]);
    const info = await readInfo(result.bytes);
    expect(info.pageCount).toBe(1);
    expect(await pageText(result.bytes, 1)).toBe('World');
  });

  it('deletes the run when newText is empty', async () => {
    const bytes = await onePageDoc(['Hello World', 'Second line']);
    const rect = await rectFor(bytes, 1, 'Hello World');
    const result = await applyTextEdits(bytes, [{ page: 1, rect, oldText: 'Hello World', newText: '' }]);
    expect(result.applied).toBe(1);
    const text = await pageText(result.bytes, 1);
    expect(text).not.toContain('Hello');
    expect(text).toContain('Second line');
  });

  it('is whitespace-insensitive when matching oldText', async () => {
    const bytes = await onePageDoc(['Hello World']);
    const rect = await rectFor(bytes, 1, 'Hello World');
    const result = await applyTextEdits(bytes, [{ page: 1, rect, oldText: '  Hello   World  ', newText: 'World' }]);
    expect(result.applied).toBe(1);
    expect(result.skipped).toEqual([]);
  });
});

describe('applyTextEdits: font-rebuild path (P1.2)', () => {
  it('rebuilds for characters outside the run\'s current subset, via the resolved system font', async () => {
    const bytes = await onePageDoc(['Hello World']);
    const rect = await rectFor(bytes, 1, 'Hello World');
    // 'z' and 'Q' are never drawn anywhere in this document's font — SetText reuse is
    // unavailable, so this must go through resolveFontBytes + FPDFText_LoadFont.
    const result = await applyTextEdits(bytes, [{ page: 1, rect, oldText: 'Hello World', newText: 'Qzzz' }]);
    expect(result.skipped).toEqual([]);
    expect(result.applied).toBe(1);
    expect(await pageText(result.bytes, 1)).toBe('Qzzz');
  });

  it('rebuilds for non-ASCII replacement text', async () => {
    const bytes = await onePageDoc(['Hello World']);
    const rect = await rectFor(bytes, 1, 'Hello World');
    const result = await applyTextEdits(bytes, [{ page: 1, rect, oldText: 'Hello World', newText: 'héllo' }]);
    expect(result.skipped).toEqual([]);
    expect(result.applied).toBe(1);
    expect(await pageText(result.bytes, 1)).toBe('héllo');
  });

  it('rebuilds a multi-line replacement as stacked text objects', async () => {
    const bytes = await onePageDoc(['Hello World']);
    const rect = await rectFor(bytes, 1, 'Hello World');
    const result = await applyTextEdits(bytes, [{ page: 1, rect, oldText: 'Hello World', newText: 'Qzzz one\nQzzz two' }]);
    expect(result.skipped).toEqual([]);
    expect(result.applied).toBe(1);
    const text = await pageText(result.bytes, 1);
    expect(text).toContain('Qzzz one');
    expect(text).toContain('Qzzz two');
  });

  it('rebuilds CJK text through the fallback face cascade (exercises real CJK system fonts)', async () => {
    const bytes = await onePageDoc(['Hello World']);
    const rect = await rectFor(bytes, 1, 'Hello World');
    // No fallback face covering CJK on this runner just means the edit is skipped, not
    // corrupted — this test only asserts consistency (applied XOR skipped), not that a
    // CJK face happens to be installed everywhere CI runs.
    const result = await applyTextEdits(bytes, [{ page: 1, rect, oldText: 'Hello World', newText: '你好世界' }]);
    expect(result.applied + result.skipped.length).toBe(1);
    if (result.applied === 1) {
      expect(await pageText(result.bytes, 1)).toBe('你好世界');
    }
  });
});

describe('applyTextEdits: fail-closed skips', () => {
  it('skips text no available font can draw, without corrupting the file', async () => {
    const bytes = await onePageDoc(['Hello World']);
    const rect = await rectFor(bytes, 1, 'Hello World');
    // A Supplementary Private Use Area codepoint: essentially guaranteed unmapped by
    // every installed font, so resolveFontBytes exhausts every candidate and throws.
    const result = await applyTextEdits(bytes, [{ page: 1, rect, oldText: 'Hello World', newText: '\u{F0000}' }]);
    expect(result.applied).toBe(0);
    expect(result.skipped).toHaveLength(1);
    expect(result.skipped[0]?.reason).toMatch(/no available font can draw/);
    // Original bytes returned untouched.
    expect(await pageText(result.bytes, 1)).toBe('Hello World');
  });

  it('skips an edit whose oldText does not match anything on the page', async () => {
    const bytes = await onePageDoc(['Hello World']);
    const result = await applyTextEdits(bytes, [{ page: 1, rect: [0, 0, 100, 20], oldText: 'Nonexistent', newText: 'World' }]);
    expect(result.applied).toBe(0);
    expect(result.skipped[0]?.reason).toMatch(/could not be located/);
  });

  it('skips an edit targeting a page that does not exist', async () => {
    const bytes = await onePageDoc(['Hello World']);
    const result = await applyTextEdits(bytes, [{ page: 5, rect: [0, 0, 100, 20], oldText: 'Hello World', newText: 'World' }]);
    expect(result.applied).toBe(0);
    expect(result.skipped[0]?.reason).toMatch(/does not exist/);
  });

  it('returns the original bytes untouched when nothing applies', async () => {
    const bytes = await onePageDoc(['Hello World']);
    const result = await applyTextEdits(bytes, [{ page: 1, rect: [0, 0, 1, 1], oldText: 'Nonexistent', newText: 'World' }]);
    expect(result.bytes).toBe(bytes);
  });

  it('two edits resolving to the same object: the second is skipped, not double-applied', async () => {
    const bytes = await onePageDoc(['Hello World']);
    const rect = await rectFor(bytes, 1, 'Hello World');
    const result = await applyTextEdits(bytes, [
      { page: 1, rect, oldText: 'Hello World', newText: 'World' },
      { page: 1, rect, oldText: 'Hello World', newText: 'Hello' },
    ]);
    expect(result.applied).toBe(1);
    expect(result.skipped).toHaveLength(1);
    expect(result.skipped[0]?.reason).toMatch(/overlaps another pending text edit/);
  });
});

describe('applyTextEdits: multi-page', () => {
  it('applies independently on each targeted page', async () => {
    const first = await onePageDoc(['Page one text']);
    const second = await onePageDoc(['Page two text']);
    const merged = await PDFDocument.create();
    for (const bytes of [first, second]) {
      const src = await PDFDocument.load(bytes);
      const [copied] = await merged.copyPages(src, [0]);
      merged.addPage(copied);
    }
    const bytes = await merged.save();
    const rect1 = await rectFor(bytes, 1, 'Page one text');
    const rect2 = await rectFor(bytes, 2, 'Page two text');
    const result = await applyTextEdits(bytes, [
      { page: 1, rect: rect1, oldText: 'Page one text', newText: 'one' },
      { page: 2, rect: rect2, oldText: 'Page two text', newText: 'two' },
    ]);
    expect(result.applied).toBe(2);
    expect(await pageText(result.bytes, 1)).toBe('one');
    expect(await pageText(result.bytes, 2)).toBe('two');
  });
});

describe('applyTextEdits: matching cascade (P1.3)', () => {
  it('splices a fragment into a single object via the container path (oldText is a substring, not the whole object)', async () => {
    const bytes = await onePageDoc(['Hello World Foo']);
    // Full object bounds as the rect: the primary whole-match path fails (oldText is
    // only part of the object's text), so this must fall through to matchEdit's
    // container/fragment branch (rectCoverage(object, rect) >= 0.5).
    const rect = await rectFor(bytes, 1, 'Hello World Foo');
    const result = await applyTextEdits(bytes, [{ page: 1, rect, oldText: 'World', newText: 'Nick' }]);
    expect(result.skipped).toEqual([]);
    expect(result.applied).toBe(1);
    expect(await pageText(result.bytes, 1)).toBe('Hello Nick Foo');
  });

  it('resolves a multi-object whole match (two separate text objects joined by the edit rect)', async () => {
    const doc = await PDFDocument.create();
    const font = await doc.embedFont(StandardFonts.Helvetica);
    const page = doc.addPage([400, 100]);
    const size = 24;
    page.drawText('Hello ', { x: 40, y: 50, size, font });
    const helloWidth = font.widthOfTextAtSize('Hello ', size);
    page.drawText('World', { x: 40 + helloWidth, y: 50, size, font });
    const bytes = await doc.save();

    const info = await readInfo(bytes);
    expect(info.pageCount).toBe(1);
    const rect: [number, number, number, number] = [40, 50, 40 + helloWidth + font.widthOfTextAtSize('World', size), 50 + size];
    // Neither object alone spans the full rect, but their JOINED text equals oldText —
    // this must resolve through the primary path's multi-object candidate set, and
    // (since matches.length > 1) always rebuild rather than SetText.
    const result = await applyTextEdits(bytes, [{ page: 1, rect, oldText: 'Hello World', newText: 'Hi there' }]);
    expect(result.skipped).toEqual([]);
    expect(result.applied).toBe(1);
    expect(await pageText(bytes, 1)).toBe('Hello World'); // sanity: original untouched
    expect(await pageText(result.bytes, 1)).toBe('Hi there');
  });

  it('preserves the untouched suffix of a container fragment edit (codepoint-preserving splice)', async () => {
    const bytes = await onePageDoc(['abc123xyz']);
    const rect = await rectFor(bytes, 1, 'abc123xyz');
    const result = await applyTextEdits(bytes, [{ page: 1, rect, oldText: '123', newText: '456' }]);
    expect(result.skipped).toEqual([]);
    expect(await pageText(result.bytes, 1)).toBe('abc456xyz');
  });
});

describe('applyTextEdits: block reflow (P1.4)', () => {
  it('positions a reflowed multi-line replacement at the given origin/lineLeading, not the anchor\'s own matrix', async () => {
    const bytes = await onePageDoc(['Hello World']);
    const rect = await rectFor(bytes, 1, 'Hello World');
    const origin: [number, number] = [40, 100];
    const result = await applyTextEdits(bytes, [
      { page: 1, rect, oldText: 'Hello World', newText: 'line one\nline two\nline three', origin, fontSize: 12, lineLeading: 20 },
    ]);
    expect(result.skipped).toEqual([]);
    expect(result.applied).toBe(1);

    const doc = await PdfViewerDocument.load(result.bytes);
    try {
      const page = await doc.getPage(1);
      const content = await page.getTextContent();
      const items = content.items.filter((it) => 'str' in it) as { str: string; transform: number[] }[];
      const lineOne = items.find((it) => it.str === 'line one');
      const lineTwo = items.find((it) => it.str === 'line two');
      const lineThree = items.find((it) => it.str === 'line three');
      expect(lineOne).toBeDefined();
      expect(lineTwo).toBeDefined();
      expect(lineThree).toBeDefined();
      // First line lands exactly at origin; each later line steps down by lineLeading.
      expect(lineOne!.transform[4]).toBeCloseTo(origin[0], 1);
      expect(lineOne!.transform[5]).toBeCloseTo(origin[1], 1);
      expect(lineTwo!.transform[5]).toBeCloseTo(origin[1] - 20, 1);
      expect(lineThree!.transform[5]).toBeCloseTo(origin[1] - 40, 1);
    } finally {
      await doc.destroy();
    }
  });

  it('applies a per-line x offset', async () => {
    const bytes = await onePageDoc(['Hello World']);
    const rect = await rectFor(bytes, 1, 'Hello World');
    const origin: [number, number] = [40, 100];
    const result = await applyTextEdits(bytes, [
      { page: 1, rect, oldText: 'Hello World', newText: 'a\nb', origin, fontSize: 12, lineLeading: 20, lineXOffsets: [0, 15] },
    ]);
    expect(result.applied).toBe(1);
    const doc = await PdfViewerDocument.load(result.bytes);
    try {
      const page = await doc.getPage(1);
      const content = await page.getTextContent();
      const items = content.items.filter((it) => 'str' in it) as { str: string; transform: number[] }[];
      const a = items.find((it) => it.str === 'a');
      const b = items.find((it) => it.str === 'b');
      expect(a!.transform[4]).toBeCloseTo(origin[0], 1);
      expect(b!.transform[4]).toBeCloseTo(origin[0] + 15, 1);
    } finally {
      await doc.destroy();
    }
  });

  it('strips position overrides for a fragment (non-whole) match instead of dragging surrounding text to the origin', async () => {
    const bytes = await onePageDoc(['abc123xyz']);
    const rect = await rectFor(bytes, 1, 'abc123xyz');
    const farOrigin: [number, number] = [350, 10]; // opposite corner from onePageDoc's [40, 250] draw position
    // origin set WITHOUT lineLeading: useMerge stays false so the container/fragment
    // path is still reachable; apply.ts must still strip origin before rebuilding
    // since whole=false here, or "abc456xyz" would jump to farOrigin instead of
    // staying where "abc123xyz" was originally drawn.
    const result = await applyTextEdits(bytes, [{ page: 1, rect, oldText: '123', newText: '456', origin: farOrigin, fontSize: 12 }]);
    expect(result.skipped).toEqual([]);
    expect(await pageText(result.bytes, 1)).toBe('abc456xyz');

    const doc = await PdfViewerDocument.load(result.bytes);
    try {
      const page = await doc.getPage(1);
      const content = await page.getTextContent();
      const item = content.items.find((it) => 'str' in it && it.str === 'abc456xyz') as { transform: number[] } | undefined;
      expect(item).toBeDefined();
      expect(item!.transform[4]).not.toBeCloseTo(farOrigin[0], 1);
      expect(item!.transform[5]).not.toBeCloseTo(farOrigin[1], 1);
    } finally {
      await doc.destroy();
    }
  });
});

describe('applyTextEdits: style overrides (P1.5)', () => {
  it('rebuilds via a curated font (newFont) even for otherwise-reusable ASCII text', async () => {
    const bytes = await onePageDoc(['Hello World']);
    const rect = await rectFor(bytes, 1, 'Hello World');
    // 'World' alone would qualify for the SetText fast path with no override — newFont
    // must force the rebuild path instead (canReuseFont rejects any style override).
    const result = await applyTextEdits(bytes, [{ page: 1, rect, oldText: 'Hello World', newText: 'World', newFont: 'times' }]);
    expect(result.skipped).toEqual([]);
    expect(result.applied).toBe(1);
    expect(await pageText(result.bytes, 1)).toBe('World');
  });

  it('rebuilds with a bold/italic toggle on the original font (no newFont)', async () => {
    const bytes = await onePageDoc(['Hello World']);
    const rect = await rectFor(bytes, 1, 'Hello World');
    const result = await applyTextEdits(bytes, [{ page: 1, rect, oldText: 'Hello World', newText: 'Hello World', newBold: true }]);
    expect(result.skipped).toEqual([]);
    expect(result.applied).toBe(1);
    expect(await pageText(result.bytes, 1)).toBe('Hello World');
  });

  it('rebuilds with a uniform color override', async () => {
    const bytes = await onePageDoc(['Hello World']);
    const rect = await rectFor(bytes, 1, 'Hello World');
    const result = await applyTextEdits(bytes, [{ page: 1, rect, oldText: 'Hello World', newText: 'World', newColor: [255, 0, 0] }]);
    expect(result.skipped).toEqual([]);
    expect(result.applied).toBe(1);
    expect(await pageText(result.bytes, 1)).toBe('World');
  });

  it('rebuilds with a font-size override', async () => {
    const bytes = await onePageDoc(['Hello World']);
    const rect = await rectFor(bytes, 1, 'Hello World');
    const result = await applyTextEdits(bytes, [{ page: 1, rect, oldText: 'Hello World', newText: 'World', fontSize: 40 }]);
    expect(result.skipped).toEqual([]);
    expect(result.applied).toBe(1);
  });

  it('style overrides still apply through the fragment (container) match path', async () => {
    const bytes = await onePageDoc(['abc123xyz']);
    const rect = await rectFor(bytes, 1, 'abc123xyz');
    const result = await applyTextEdits(bytes, [{ page: 1, rect, oldText: '123', newText: '456', newColor: [0, 128, 0] }]);
    expect(result.skipped).toEqual([]);
    expect(await pageText(result.bytes, 1)).toBe('abc456xyz');
  });

  it('a uniform newColor override actually lands as a real fill-color operator', async () => {
    const bytes = await onePageDoc(['Hello World']);
    const rect = await rectFor(bytes, 1, 'Hello World');
    const result = await applyTextEdits(bytes, [{ page: 1, rect, oldText: 'Hello World', newText: 'World', newColor: [255, 0, 0] }]);
    const colors = distinctFillColors(await pageContentStreamText(result.bytes, 1));
    expect(colors.has('1,0,0')).toBe(true);
  });
});

describe('applyTextEdits: per-character style runs (P1.6)', () => {
  it('splits a whole-match replacement into differently-colored segments', async () => {
    const bytes = await onePageDoc(['Hello World']);
    const rect = await rectFor(bytes, 1, 'Hello World');
    const result = await applyTextEdits(bytes, [
      {
        page: 1,
        rect,
        oldText: 'Hello World',
        newText: 'red green',
        styleRuns: [
          { start: 0, end: 3, color: [255, 0, 0] },
          { start: 4, end: 9, color: [0, 128, 0] },
        ],
      },
    ]);
    expect(result.skipped).toEqual([]);
    expect(result.applied).toBe(1);
    expect(await pageText(result.bytes, 1)).toBe('red green');
    const colors = distinctFillColors(await pageContentStreamText(result.bytes, 1));
    expect(colors.has('1,0,0')).toBe(true);
    // Two distinct styled colors actually landed as two distinct fill-color operators
    // (exact decimal formatting of 128/255 is pdfium's own float precision, not ours to pin).
    expect(colors.size).toBeGreaterThanOrEqual(2);
  });

  it('leaves an unstyled prefix/suffix in the base color while styling only the run', async () => {
    const bytes = await onePageDoc(['Hello World']);
    const rect = await rectFor(bytes, 1, 'Hello World');
    const result = await applyTextEdits(bytes, [
      { page: 1, rect, oldText: 'Hello World', newText: 'plain RED plain', styleRuns: [{ start: 6, end: 9, color: [255, 0, 0] }] },
    ]);
    expect(result.skipped).toEqual([]);
    expect(await pageText(result.bytes, 1)).toBe('plain RED plain');
    const colors = distinctFillColors(await pageContentStreamText(result.bytes, 1));
    // Base (black, the default fill on a fresh pdf-lib page) AND the styled red both present.
    expect(colors.has('1,0,0')).toBe(true);
    expect(colors.size).toBeGreaterThanOrEqual(2);
  });

  it('aligns styleRuns (indexed into the raw newText) correctly through a fragment splice', async () => {
    // A container-path fragment: resolvedText ("abc456xyz") differs from the user's
    // raw newText ("456") that styleRuns' indices refer to — plannedCharStyles must
    // realign through the fold, not use raw offsets directly.
    const bytes = await onePageDoc(['abc123xyz']);
    const rect = await rectFor(bytes, 1, 'abc123xyz');
    const result = await applyTextEdits(bytes, [{ page: 1, rect, oldText: '123', newText: '456', styleRuns: [{ start: 0, end: 1, color: [255, 0, 0] }] }]);
    expect(result.skipped).toEqual([]);
    expect(await pageText(result.bytes, 1)).toBe('abc456xyz');
    const colors = distinctFillColors(await pageContentStreamText(result.bytes, 1));
    expect(colors.has('1,0,0')).toBe(true);
  });

  it('forces a rebuild (never the SetText fast path) even for plain reusable ASCII', async () => {
    const bytes = await onePageDoc(['Hello World']);
    const rect = await rectFor(bytes, 1, 'Hello World');
    // 'World' alone would qualify for SetText reuse with no styleRuns.
    const result = await applyTextEdits(bytes, [{ page: 1, rect, oldText: 'Hello World', newText: 'World', styleRuns: [{ start: 0, end: 2, color: [255, 0, 0] }] }]);
    expect(result.skipped).toEqual([]);
    expect(await pageText(result.bytes, 1)).toBe('World');
    const colors = distinctFillColors(await pageContentStreamText(result.bytes, 1));
    expect(colors.has('1,0,0')).toBe(true);
  });

  it('applies a bold/italic style run using a different loaded font than the base', async () => {
    const bytes = await onePageDoc(['Hello World']);
    const rect = await rectFor(bytes, 1, 'Hello World');
    const result = await applyTextEdits(bytes, [
      { page: 1, rect, oldText: 'Hello World', newText: 'plain BOLD', styleRuns: [{ start: 6, end: 10, font: 'arial', bold: true }] },
    ]);
    expect(result.skipped).toEqual([]);
    // pdf.js's text layer inserts a spurious extra space at a font-boundary between two
    // adjacent text objects (same class of pdf.js-vs-pdfium discrepancy norm() exists
    // for elsewhere in this package) — whitespace-insensitive compare, not exact.
    expect((await pageText(result.bytes, 1)).replace(/\s+/g, ' ')).toBe('plain BOLD');
  });
});

describe('applyTextEdits: character-preserving redraw (P1.7 keep-plan)', () => {
  async function twoObjectDoc(first: string, second: string): Promise<Uint8Array> {
    const doc = await PDFDocument.create();
    const font = await doc.embedFont(StandardFonts.Helvetica);
    const page = doc.addPage([400, 100]);
    const size = 24;
    page.drawText(first, { x: 40, y: 50, size, font });
    const firstWidth = font.widthOfTextAtSize(first, size);
    page.drawText(second, { x: 40 + firstWidth, y: 50, size, font });
    return doc.save();
  }

  it('keeps (byte-identically re-uses) the unchanged object instead of redrawing it', async () => {
    const bytes = await twoObjectDoc('Hello ', 'World');
    const originalHexPayloads = hexTjPayloads(await pageContentStreamText(bytes, 1));
    expect(originalHexPayloads.length).toBe(2); // one Tj per drawText call

    const rect: [number, number, number, number] = [0, 0, 400, 100]; // whole page — covers both objects
    const result = await applyTextEdits(bytes, [{ page: 1, rect, oldText: 'Hello World', newText: 'Hello Nick' }]);
    expect(result.skipped).toEqual([]);
    expect(result.applied).toBe(1);
    // pdf.js's text layer inserts a spurious extra space at the boundary between the
    // kept "Hello " object and the newly-redrawn "Nick" object (same pdf.js-vs-pdfium
    // discrepancy documented on the style-run tests) — whitespace-insensitive compare.
    expect((await pageText(result.bytes, 1)).replace(/\s+/g, ' ')).toBe('Hello Nick');

    // The "Hello " object's original glyph encoding must survive byte-identical — a
    // redraw would re-subset the font and very likely produce different glyph ids
    // even for visually-identical text. Its match rect covers only the untouched
    // prefix here, so we know exactly which original payload to look for.
    const helloOnlyDoc = await PDFDocument.create();
    const helloFont = await helloOnlyDoc.embedFont(StandardFonts.Helvetica);
    helloOnlyDoc.addPage([400, 100]).drawText('Hello ', { x: 40, y: 50, size: 24, font: helloFont });
    const helloPayload = hexTjPayloads(await pageContentStreamText(await helloOnlyDoc.save(), 1))[0]!;

    const outputPayloads = hexTjPayloads(await pageContentStreamText(result.bytes, 1));
    expect(outputPayloads).toContain(helloPayload);
  });

  it('does not accidentally keep a whitespace-boundary edit (wsEditClamp)', async () => {
    // Deleting the gap in "phon e" -> "phone" changes no NFKC-folded unit at all; a
    // naive keep-plan would keep BOTH objects verbatim at their original spacing,
    // silently undoing the edit. wsEditClamp must force the seam-adjacent glyph to
    // redraw so the gap actually closes.
    const bytes = await twoObjectDoc('phon ', 'e');
    const rect: [number, number, number, number] = [0, 0, 400, 100];
    const result = await applyTextEdits(bytes, [{ page: 1, rect, oldText: 'phon e', newText: 'phone' }]);
    expect(result.skipped).toEqual([]);
    expect(await pageText(result.bytes, 1)).toBe('phone');
  });

  it('a color-only override does not preclude keeping (only face/size overrides do)', async () => {
    const bytes = await twoObjectDoc('Hello ', 'World');
    const rect: [number, number, number, number] = [0, 0, 400, 100];
    const result = await applyTextEdits(bytes, [{ page: 1, rect, oldText: 'Hello World', newText: 'Hello Nick', newColor: [255, 0, 0] }]);
    expect(result.skipped).toEqual([]);
    // pdf.js text-boundary artifact (see the preceding test) — whitespace-insensitive.
    expect((await pageText(result.bytes, 1)).replace(/\s+/g, ' ')).toBe('Hello Nick');
    // The kept "Hello " object gets recolored via a direct SetFillColor on the
    // ORIGINAL object (buildPreserved's `moves` path), not a redraw — still real red.
    const colors = distinctFillColors(await pageContentStreamText(result.bytes, 1));
    expect(colors.has('1,0,0')).toBe(true);
  });

  it('falls back to a full redraw (never errors) for a rotated, axis-misaligned matrix', async () => {
    const doc = await PDFDocument.create();
    const font = await doc.embedFont(StandardFonts.Helvetica);
    const page = doc.addPage([400, 400]);
    // A genuinely rotated text matrix (45°) fails buildKeepPlan's axis-aligned gate.
    page.drawText('Hello World', { x: 100, y: 100, size: 24, font, rotate: degrees(45) });
    const bytes = await doc.save();
    // Full-page rect: pdfjs's transform-based bbox approximation (rectFor) isn't
    // reliable for rotated glyphs, and isn't the point of this test anyway.
    const result = await applyTextEdits(bytes, [{ page: 1, rect: [0, 0, 400, 400], oldText: 'Hello World', newText: 'Hello Nick' }]);
    expect(result.skipped).toEqual([]);
    expect(await pageText(result.bytes, 1)).toBe('Hello Nick');
  });
});

describe('editText / editBlock (P1.8 — separate public primitives)', () => {
  it('editText performs a surgical replacement identically to applyTextEdits', async () => {
    const bytes = await onePageDoc(['Hello World']);
    const rect = await rectFor(bytes, 1, 'Hello World');
    const spec: SurgicalEditSpec = { page: 1, rect, oldText: 'Hello World', newText: 'Hello Nick' };
    const result = await editText(bytes, [spec]);
    expect(result.skipped).toEqual([]);
    expect(result.applied).toBe(1);
    expect(await pageText(result.bytes, 1)).toBe('Hello Nick');
  });

  it('editBlock reflows a multi-line replacement at the required origin', async () => {
    const bytes = await onePageDoc(['Hello World']);
    const rect = await rectFor(bytes, 1, 'Hello World');
    const origin: [number, number] = [40, 100];
    const spec: BlockEditSpec = { page: 1, rect, oldText: 'Hello World', newText: 'line one\nline two', origin, fontSize: 12, lineLeading: 20 };
    const result = await editBlock(bytes, [spec]);
    expect(result.skipped).toEqual([]);
    expect(result.applied).toBe(1);

    const doc = await PdfViewerDocument.load(result.bytes);
    try {
      const page = await doc.getPage(1);
      const content = await page.getTextContent();
      const items = content.items.filter((it) => 'str' in it) as { str: string; transform: number[] }[];
      const lineOne = items.find((it) => it.str === 'line one');
      expect(lineOne).toBeDefined();
      expect(lineOne!.transform[4]).toBeCloseTo(origin[0], 1);
      expect(lineOne!.transform[5]).toBeCloseTo(origin[1], 1);
    } finally {
      await doc.destroy();
    }
  });

  it('SurgicalEditSpec rejects origin/lineLeading/lineXOffsets at compile time', () => {
    // @ts-expect-error — origin is not a field of SurgicalEditSpec (that's the whole point).
    const withOrigin: SurgicalEditSpec = { page: 1, rect: [0, 0, 1, 1], oldText: 'a', newText: 'b', origin: [0, 0] };
    expect(withOrigin).toBeDefined();
  });

  it('BlockEditSpec requires origin at compile time', () => {
    // @ts-expect-error — origin is required on BlockEditSpec, omitting it is a type error.
    const withoutOrigin: BlockEditSpec = { page: 1, rect: [0, 0, 1, 1], oldText: 'a', newText: 'b' };
    expect(withoutOrigin).toBeDefined();
  });
});
