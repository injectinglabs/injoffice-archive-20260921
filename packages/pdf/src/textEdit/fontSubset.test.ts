import { existsSync, readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { identityCffCharset, subsetTtf } from './fontSubset.js';

// Real system fonts, same posture as font-metrics' own tests: no synthetic
// fixtures for something whose whole point is parsing real-world files.
const TTF_CANDIDATES = ['/System/Library/Fonts/SFNSMono.ttf', '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf'];
const OTF_CANDIDATES = ['/System/Library/Fonts/Supplemental/STIXGeneral.otf'];

function firstExisting(paths: string[]): string | null {
  return paths.find((p) => existsSync(p)) ?? null;
}

const isSfntStart = (buf: Buffer) => {
  const tag = buf.readUInt32BE(0);
  return tag === 0x00010000 || tag === 0x74727565 || tag === 0x4f54544f;
};

describe('subsetTtf', () => {
  it('subsets a real TrueType font down to a small valid sfnt', async () => {
    const path = firstExisting(TTF_CANDIDATES);
    if (!path) return; // no Linux/macOS TTF present on this runner
    const font = readFileSync(path);
    const subset = await subsetTtf(font, 'Hello');
    expect(isSfntStart(subset)).toBe(true);
    expect(subset.length).toBeGreaterThan(0);
    expect(subset.length).toBeLessThan(font.length);
  });

  it('subsets a real CFF-flavored OpenType font', async () => {
    const path = firstExisting(OTF_CANDIDATES);
    if (!path) return;
    const font = readFileSync(path);
    const subset = await subsetTtf(font, 'ABC');
    expect(isSfntStart(subset)).toBe(true);
    expect(subset.length).toBeGreaterThan(0);
  });

  it('rejects a garbage buffer', async () => {
    await expect(subsetTtf(Buffer.from('not a font'), 'x')).rejects.toThrow();
  });
});

describe('identityCffCharset', () => {
  it('passes a non-OTTO (TrueType) font through unchanged', () => {
    const path = firstExisting(TTF_CANDIDATES);
    if (!path) return;
    const font = readFileSync(path);
    expect(identityCffCharset(font)).toBe(font);
  });

  it('does not throw on a real CFF font (rewrites if CID-keyed, passes through otherwise)', () => {
    const path = firstExisting(OTF_CANDIDATES);
    if (!path) return;
    const font = readFileSync(path);
    const out = identityCffCharset(font);
    expect(out.length).toBeGreaterThan(0);
    expect(out.readUInt32BE(0)).toBe(0x4f54544f);
  });

  it('passes a too-short buffer through unchanged rather than throwing', () => {
    const tiny = Buffer.from([1, 2, 3]);
    expect(identityCffCharset(tiny)).toBe(tiny);
  });
});
