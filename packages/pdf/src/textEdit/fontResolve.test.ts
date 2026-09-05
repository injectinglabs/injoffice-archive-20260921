import { describe, expect, it } from 'vitest';
import { canDrawText, listEditFonts } from './fontResolve.js';

describe('listEditFonts', () => {
  it('finds at least one curated font readable on this machine', () => {
    expect(listEditFonts().length).toBeGreaterThan(0);
  });

  it('only returns known curated ids', () => {
    for (const id of listEditFonts()) expect(['arial', 'times', 'courier']).toContain(id);
  });
});

describe('canDrawText', () => {
  it('is true for plain ASCII via the fallback cascade with no style request', () => {
    expect(canDrawText('Hello World')).toBe(true);
  });

  it('is true for a curated font that resolves on this machine', () => {
    const [available] = listEditFonts();
    if (!available) return; // no curated font readable here — nothing to assert
    expect(canDrawText('Hello', { font: available })).toBe(true);
  });

  it('is true for a curated font + bold + italic request', () => {
    const [available] = listEditFonts();
    if (!available) return;
    expect(canDrawText('Hello', { font: available, bold: true, italic: true })).toBe(true);
  });
});
