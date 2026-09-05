import { describe, expect, it } from 'vitest';
import { canReuseFont, resolveMatch, type PageTextObj } from './match.js';

// Synthetic PageTextObj fixtures — this is the matching LOGIC in isolation, no real
// pdfium engine needed, which lets rect geometry be picked exactly to hit each branch
// of the cascade deterministically (something real-engine-driven tests in apply.test.ts
// can't easily guarantee for the harder rescue paths).

function obj(index: number, text: string, bounds: [number, number, number, number], font = 1): PageTextObj {
  return { obj: index + 1, index, text, font, bounds };
}

describe('resolveMatch: primary whole-match', () => {
  it('matches a single object whose text equals oldText', () => {
    const objects = [obj(0, 'Hello World', [0, 0, 100, 20])];
    const result = resolveMatch(objects, [0, 0, 100, 20], 'Hello World', 'Hi there');
    expect('reason' in result).toBe(false);
    if ('reason' in result) return;
    expect(result.matches).toEqual([objects[0]]);
    expect(result.whole).toBe(true);
    expect(result.newText).toBe('Hi there');
  });

  it('matches two adjacent objects joined in stream order', () => {
    const a = obj(0, 'Hello ', [0, 0, 50, 20]);
    const b = obj(1, 'World', [50, 0, 90, 20]);
    const result = resolveMatch([a, b], [0, 0, 90, 20], 'Hello World', 'Hi there');
    expect('reason' in result).toBe(false);
    if ('reason' in result) return;
    expect(result.matches).toEqual([a, b]);
    expect(result.whole).toBe(true);
  });

  it('recovers a stacked xLayer when two runs overlap the same rect (overflow scenario)', () => {
    // Two full-line runs stacked at the same bounds (an earlier overflowing edit left
    // two objects drawn over each other) — stream order alone won't equal either line,
    // but the x-layer split isolates each one.
    const line1 = obj(0, 'first line', [0, 0, 100, 20]);
    const line2 = obj(1, 'second line', [0, 0, 100, 20]);
    const result = resolveMatch([line1, line2], [0, 0, 100, 20], 'second line', 'replaced');
    expect('reason' in result).toBe(false);
    if ('reason' in result) return;
    expect(result.matches).toEqual([line2]);
    expect(result.whole).toBe(true);
  });
});

describe('resolveMatch: container/fragment path', () => {
  it('splices a fragment into the single object that covers most of the rect', () => {
    const objects = [obj(0, 'abc123xyz', [0, 0, 100, 20])];
    const result = resolveMatch(objects, [0, 0, 100, 20], '123', '456');
    expect('reason' in result).toBe(false);
    if ('reason' in result) return;
    expect(result.matches).toEqual([objects[0]]);
    expect(result.whole).toBe(false);
    expect(result.newText).toBe('abc456xyz');
  });

  it('does not match when the fragment text is not present in the covering object', () => {
    const objects = [obj(0, 'abcxyz', [0, 0, 100, 20])];
    const result = resolveMatch(objects, [0, 0, 100, 20], '123', '456');
    expect('reason' in result).toBe(true);
  });
});

describe('resolveMatch: matchByText rescue', () => {
  // Two objects with a real gap between them (representing inter-word spacing) — a
  // rect straddling the gap fails BOTH overlapRatio>=0.5 (against each object's own
  // area) and rectCoverage>=0.5 (against the rect's own area), so only the padded
  // touch-search in matchByText can find them.
  const a = obj(0, 'Hello ', [0, 0, 50, 20]);
  const b = obj(1, 'World', [56, 0, 96, 20]);
  const gapRect: [number, number, number, number] = [48, 0, 58, 20];

  it('does not resolve via the primary or container path for a rect that undershoots both objects', () => {
    // Sanity check on the geometry itself, independent of the rescue: neither branch
    // ahead of matchByText should be satisfied by gapRect.
    const primaryOverlap = (o: PageTextObj) => {
      const w = Math.min(o.bounds[2], gapRect[2]) - Math.max(o.bounds[0], gapRect[0]);
      const h = Math.min(o.bounds[3], gapRect[3]) - Math.max(o.bounds[1], gapRect[1]);
      const area = w > 0 && h > 0 ? w * h : 0;
      const objArea = (o.bounds[2] - o.bounds[0]) * (o.bounds[3] - o.bounds[1]);
      return area / objArea;
    };
    expect(primaryOverlap(a)).toBeLessThan(0.5);
    expect(primaryOverlap(b)).toBeLessThan(0.5);
  });

  it('rescues a whole match spanning both objects via the padded touch search', () => {
    const result = resolveMatch([a, b], gapRect, 'Hello World', 'Hi there');
    expect('reason' in result).toBe(false);
    if ('reason' in result) return;
    expect(result.matches).toEqual([a, b]);
    expect(result.whole).toBe(true);
    expect(result.newText).toBe('Hi there');
  });

  it('rescues a fragment inside one of the touched objects when oldText is only part of it', () => {
    const result = resolveMatch([a, b], gapRect, 'Wor', 'xyz');
    expect('reason' in result).toBe(false);
    if ('reason' in result) return;
    expect(result.matches).toEqual([b]);
    expect(result.whole).toBe(false);
    expect(result.newText).toBe('xyzld');
  });

  it('refuses (does not degrade the layout) a rescued FRAGMENT when useMerge (lineLeading) is set', () => {
    const wholeResult = resolveMatch([a, b], gapRect, 'Hello World', 'Hi there', true);
    expect('reason' in wholeResult).toBe(false); // a whole rescue is still fine with useMerge

    const fragmentResult = resolveMatch([a, b], gapRect, 'Wor', 'xyz', true);
    expect('reason' in fragmentResult).toBe(true);
  });

  it('returns a not-found reason when nothing in the padded touch zone contains oldText', () => {
    const result = resolveMatch([a, b], gapRect, 'Nonexistent', 'xyz');
    expect('reason' in result).toBe(true);
  });
});

describe('resolveMatch: deletion normalization', () => {
  it('normalizes a whitespace-only replacement to empty (deletion)', () => {
    const objects = [obj(0, 'Hello World', [0, 0, 100, 20])];
    const result = resolveMatch(objects, [0, 0, 100, 20], 'Hello World', '   ');
    expect('reason' in result).toBe(false);
    if ('reason' in result) return;
    expect(result.newText).toBe('');
  });
});

describe('canReuseFont', () => {
  const shared = 5;
  const all: PageTextObj[] = [obj(0, 'Hello World', [0, 0, 100, 20], shared)];

  it('is false for a multi-object match regardless of text', () => {
    const multi = [all[0]!, obj(1, 'x', [0, 0, 1, 1], shared)];
    expect(canReuseFont('World', multi, all)).toBe(false);
  });

  it('is false for non-ASCII text', () => {
    expect(canReuseFont('café', [all[0]!], all)).toBe(false);
  });

  it('is false when a character is not already drawn somewhere in the same font', () => {
    expect(canReuseFont('Qzzz', [all[0]!], all)).toBe(false);
  });

  it('is true when every non-space character is already drawn in the same font', () => {
    expect(canReuseFont('World', [all[0]!], all)).toBe(true);
  });
});
