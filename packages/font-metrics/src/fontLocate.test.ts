import { describe, expect, it } from 'vitest';
import { findFontCovering, findSystemFont, isFamilyInstalled, isTruetype } from './fontLocate.js';
import { getFontIndex, norm, resetFontIndexForTests, styleScore, styleTokens } from './sfnt.js';

// Host integration: walk real OS font directories (macOS Library paths,
// Linux /usr/share/fonts, Windows %WINDIR%/Fonts) and parse installed
// sfnt / ttc files. No synthetic fixtures — coverage is "does this read
// the fonts actually present on this machine".

function isSfntStart(buf: Buffer): boolean {
  const tag = buf.readUInt32BE(0);
  return tag === 0x00010000 || tag === 0x74727565 || tag === 0x4f54544f; // 1.0 / 'true' / 'OTTO'
}

describe('getFontIndex', () => {
  it('finds at least one installed font family on this machine', { timeout: 30_000 }, () => {
    resetFontIndexForTests();
    const index = getFontIndex();
    expect(index.byFamily.size).toBeGreaterThan(0);
  });
});

describe('isFamilyInstalled', () => {
  it('is true for a family this machine actually has', () => {
    const [anyFamily] = getFontIndex().byFamily.keys();
    expect(anyFamily).toBeDefined();
    expect(isFamilyInstalled(anyFamily!)).toBe(true);
  });

  it('matches the same family regardless of letter case', () => {
    const [anyFamily] = getFontIndex().byFamily.keys();
    expect(isFamilyInstalled(anyFamily!.toUpperCase())).toBe(true);
  });

  it('is false for a family that does not exist', () => {
    expect(isFamilyInstalled('Definitely Not A Real Font Family 12345')).toBe(false);
  });

  it('is false for an empty family', () => {
    expect(isFamilyInstalled('')).toBe(false);
  });
});

describe('findSystemFont', () => {
  it('resolves a real installed PostScript name to valid sfnt bytes', () => {
    const index = getFontIndex();
    let found: Buffer | null = null;
    for (const psKey of index.byPs.keys()) {
      found = findSystemFont(psKey, '');
      if (found) break;
    }
    expect(found).not.toBeNull();
    expect(isSfntStart(found!)).toBe(true);
  });

  it('resolves an installed family when the PostScript name is empty', () => {
    const [family] = getFontIndex().byFamily.keys();
    expect(family).toBeDefined();
    const bytes = findSystemFont('', family!);
    expect(bytes).not.toBeNull();
    expect(isSfntStart(bytes!)).toBe(true);
  });

  it('returns a standalone sfnt header when the matched face lives in a .ttc', () => {
    const index = getFontIndex();
    let psKey: string | undefined;
    for (const [key, face] of index.byPs) {
      if (face.offset > 0) {
        psKey = key;
        break;
      }
    }
    if (!psKey) {
      // Linux hosts often ship only standalone ttf/otf files.
      expect(index.byFamily.size).toBeGreaterThan(0);
      return;
    }
    const bytes = findSystemFont(psKey, '');
    expect(bytes).not.toBeNull();
    expect(bytes!.readUInt32BE(0)).not.toBe(0x74746366);
    expect(isSfntStart(bytes!)).toBe(true);
  });

  it('returns null when neither name matches anything installed', () => {
    expect(findSystemFont('NoSuchPostScriptName123', 'NoSuchFamily123')).toBeNull();
  });
});

describe('findFontCovering', () => {
  it('finds a face covering plain ASCII text', { timeout: 30_000 }, () => {
    const bytes = findFontCovering('Hello, world!');
    expect(bytes).not.toBeNull();
    expect(isSfntStart(bytes!)).toBe(true);
  });

  it('returns null for an empty string', () => {
    expect(findFontCovering('')).toBeNull();
  });

  it('returns null when only newlines remain', () => {
    expect(findFontCovering('\n\r\n')).toBeNull();
  });
});

describe('isTruetype', () => {
  it('is false for an OTTO (CFF-flavored) header', () => {
    const b = Buffer.alloc(4);
    b.writeUInt32BE(0x4f54544f, 0);
    expect(isTruetype(b)).toBe(false);
  });

  it('is true for a glyf-flavored (0x00010000) header', () => {
    const b = Buffer.alloc(4);
    b.writeUInt32BE(0x00010000, 0);
    expect(isTruetype(b)).toBe(true);
  });

  it('is true for an Apple TrueType (true) header', () => {
    const b = Buffer.alloc(4);
    b.writeUInt32BE(0x74727565, 0);
    expect(isTruetype(b)).toBe(true);
  });

  it('is false for a too-short buffer', () => {
    expect(isTruetype(Buffer.alloc(2))).toBe(false);
  });
});

describe('norm / styleTokens / styleScore', () => {
  it('norm folds case, whitespace, and separators', () => {
    expect(norm('Times New-Roman_Bold')).toBe('timesnewromanbold');
  });

  it('styleTokens splits BoldItalic into two independent tokens', () => {
    expect(styleTokens('BoldItalic')).toEqual(['bold', 'italic']);
  });

  it('styleTokens folds oblique into italic', () => {
    expect(styleTokens('Oblique')).toEqual(['italic']);
  });

  it('styleTokens keeps semibold distinct from bold', () => {
    expect(styleTokens('Semibold')).toEqual(['semibold']);
  });

  it('styleScore prefers an exact single-style match over a multi-style face', () => {
    const bold = { path: '', offset: 0, style: 'bold' };
    const boldItalic = { path: '', offset: 0, style: 'bold italic' };
    const scoreForBoldWant = (f: typeof bold) => styleScore(f, ['bold']);
    expect(scoreForBoldWant(bold)).toBeGreaterThan(scoreForBoldWant(boldItalic));
  });
});
