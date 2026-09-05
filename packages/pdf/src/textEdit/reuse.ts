import type { TextObjectStyle } from './styleRuns.js';

export type ReuseDecision =
  | { readonly reusable: true }
  | { readonly reusable: false; readonly reason: string };

export function canReuseTextObject(style: TextObjectStyle, replacement: string): ReuseDecision {
  if (replacement.length === 0) return { reusable: false, reason: 'PDFium cannot assign an empty string to a text object' };
  if (!Number.isFinite(style.fontSize) || style.fontSize <= 0) return { reusable: false, reason: 'source text object has an invalid font size' };
  if (![style.matrix.a, style.matrix.b, style.matrix.c, style.matrix.d, style.matrix.e, style.matrix.f].every(Number.isFinite)) {
    return { reusable: false, reason: 'source text object has a non-finite transform' };
  }
  return { reusable: true };
}
