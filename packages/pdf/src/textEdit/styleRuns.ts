import type { PdfiumSession } from './pdfium.js';

export interface TextMatrix {
  readonly a: number;
  readonly b: number;
  readonly c: number;
  readonly d: number;
  readonly e: number;
  readonly f: number;
}

export interface TextObjectStyle {
  readonly font: number;
  readonly fontSize: number;
  readonly fill: readonly [number, number, number, number];
  readonly renderMode: number;
  readonly matrix: TextMatrix;
}

export function readTextObjectStyle(session: PdfiumSession, object: number): TextObjectStyle {
  return session.readTextObjectStyle(object);
}
