import { describe, expect, it } from 'vitest';
import { foldRadicals, RADICAL_EQUIV } from './radicals.js';

// Codepoints referenced by \u escape, not typed as literal glyphs — this table is
// exactly the kind of data where a transcription slip between visually-similar CJK
// characters would silently pass a wrong test. Expected values are read directly
// out of RADICAL_EQUIV (the ported table itself), never retyped by hand.

describe('foldRadicals', () => {
  it('folds a Kangxi radical to its unified ideograph', () => {
    const kangxiPerson = '⼈'; // KANGXI RADICAL PERSON
    expect(RADICAL_EQUIV[kangxiPerson]).toBe('人'); // 'person'
    expect(foldRadicals(kangxiPerson)).toBe(RADICAL_EQUIV[kangxiPerson]);
  });

  it('folds a Radicals Supplement codepoint (no NFKC decomposition exists for it)', () => {
    const radicalsSupplementPerson = '⺅'; // CJK RADICAL PERSON
    expect(RADICAL_EQUIV[radicalsSupplementPerson]).toBeDefined();
    expect(foldRadicals(radicalsSupplementPerson)).toBe(RADICAL_EQUIV[radicalsSupplementPerson]);
    // NFKC alone leaves this block untouched — that's the whole reason this table exists.
    expect(radicalsSupplementPerson.normalize('NFKC')).toBe(radicalsSupplementPerson);
  });

  it('leaves non-radical characters untouched', () => {
    expect(foldRadicals('Hello world')).toBe('Hello world');
  });

  it('folds only the radical codepoints within a mixed string', () => {
    const kangxiPerson = '⼈';
    expect(foldRadicals(`a${kangxiPerson}b`)).toBe(`a${RADICAL_EQUIV[kangxiPerson]}b`);
  });

  it('is a no-op on an empty string', () => {
    expect(foldRadicals('')).toBe('');
  });

  it('folds every table entry to exactly its own mapped value (full round-trip)', () => {
    for (const [radical, unified] of Object.entries(RADICAL_EQUIV)) {
      expect(foldRadicals(radical)).toBe(unified);
    }
  });
});
