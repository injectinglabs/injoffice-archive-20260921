import { describe, expect, it } from 'vitest';
import { foldMap, indexOfUnits, mergeEngineCodepoints, spliceIntoEngine } from './splice.js';

describe('foldMap', () => {
  it('maps each non-space unit back to its raw char span', () => {
    const fm = foldMap('abc');
    expect(fm.units).toEqual(['a', 'b', 'c']);
    expect(fm.idx).toEqual([0, 1, 2, 3]);
    expect(fm.end).toEqual([1, 2, 3]);
  });

  it('drops whitespace by default', () => {
    expect(foldMap('a b').units).toEqual(['a', 'b']);
  });

  it('collapses a whitespace run to one unit when keepSpaces is set', () => {
    expect(foldMap('a   b', true).units).toEqual(['a', ' ', 'b']);
  });
});

describe('indexOfUnits', () => {
  it('finds a subsequence', () => {
    expect(indexOfUnits(['a', 'b', 'c', 'd'], ['b', 'c'])).toBe(1);
  });

  it('returns -1 when absent', () => {
    expect(indexOfUnits(['a', 'b'], ['x'])).toBe(-1);
  });

  it('returns -1 for an empty needle', () => {
    expect(indexOfUnits(['a', 'b'], [])).toBe(-1);
  });
});

describe('spliceIntoEngine', () => {
  it('is a pure passthrough when old and new text are identical', () => {
    expect(spliceIntoEngine('Hello World', 'Hello World', 'Hello World')).toBe('Hello World');
  });

  it('keeps the engine\'s codepoints for an unchanged prefix and suffix, typed middle only', () => {
    // Engine text uses a real space; user typed text with the same visible content.
    const out = spliceIntoEngine('Hello World', 'Hello World', 'Hello Nick');
    expect(out).toBe('Hello Nick');
  });

  it('handles a pure append at the end', () => {
    const out = spliceIntoEngine('Hello', 'Hello', 'Hello!');
    expect(out).toBe('Hello!');
  });

  it('handles a pure prepend at the start', () => {
    const out = spliceIntoEngine('World', 'World', 'Hello World');
    expect(out).toBe('Hello World');
  });

  it('handles a full replacement with nothing in common', () => {
    const out = spliceIntoEngine('abc', 'abc', 'xyz');
    expect(out).toBe('xyz');
  });

  it('is whitespace-insensitive when locating the common prefix/suffix', () => {
    // Engine text has different (but equivalent) whitespace than what the user typed as oldText.
    const out = spliceIntoEngine('Hello  World', 'Hello World', 'Hello there World');
    expect(out.replace(/\s+/g, ' ')).toBe('Hello there World');
  });
});

describe('mergeEngineCodepoints', () => {
  it('is a pure passthrough when nothing changed', () => {
    expect(mergeEngineCodepoints('Hello World', 'Hello World', 'Hello World')).toBe('Hello World');
  });

  it('keeps the engine codepoints for an unchanged prefix, typed text for the rest', () => {
    const out = mergeEngineCodepoints('Hello World', 'Hello World', 'Hello Nick');
    expect(out).toBe('Hello Nick');
  });

  it('falls back to newText verbatim when the engine text does not fold to the same unit count as oldText', () => {
    // eng.units.length !== oldU.units.length triggers the documented fallback.
    const out = mergeEngineCodepoints('ab', 'abc', 'xyz');
    expect(out).toBe('xyz');
  });
});
