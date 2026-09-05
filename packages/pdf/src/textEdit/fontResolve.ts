import type { PdfiumSession } from './pdfium.js';
import { prepareFont } from './fontSubset.js';

export type StandardPdfFont =
  | 'Courier' | 'Courier-Bold' | 'Courier-Oblique' | 'Courier-BoldOblique'
  | 'Helvetica' | 'Helvetica-Bold' | 'Helvetica-Oblique' | 'Helvetica-BoldOblique'
  | 'Times-Roman' | 'Times-Bold' | 'Times-Italic' | 'Times-BoldItalic'
  | 'Symbol' | 'ZapfDingbats';

export interface FontRequest {
  readonly fontBytes?: Uint8Array;
  readonly fontFaceIndex?: number;
  readonly standardFont?: StandardPdfFont;
}

export interface ResolvedFont {
  readonly handle: number;
  close(): void;
}

export function loadRequestedFont(session: PdfiumSession, text: string, request: FontRequest): ResolvedFont {
  if (request.fontBytes) {
    const prepared = prepareFont(request.fontBytes, text, request.fontFaceIndex ?? 0);
    if (prepared.faceIndex !== 0) throw new RangeError('PDFium v1 custom font loading supports only face index 0');
    const handle = session.withBytes(prepared.bytes, (pointer) =>
      session.api.FPDFText_LoadFont(session.document, pointer, prepared.bytes.byteLength, 2, true),
    );
    if (!handle) throw new Error('PDFium rejected the supplied TrueType font');
    return onceClose(handle, () => session.api.FPDFFont_Close(handle));
  }
  const name = request.standardFont ?? 'Helvetica';
  if (!base14Covers(name, text)) throw new RangeError(`${name} is restricted to printable ASCII in the v1 insertion policy`);
  const handle = session.api.FPDFText_LoadStandardFont(session.document, name);
  if (!handle) throw new Error(`PDFium could not load standard font ${name}`);
  return onceClose(handle, () => session.api.FPDFFont_Close(handle));
}

export function borrowedFont(handle: number): ResolvedFont {
  if (!handle) throw new Error('source text object has no reusable font');
  return { handle, close() {} };
}

function onceClose(handle: number, release: () => void): ResolvedFont {
  let open = true;
  return { handle, close() { if (open) { open = false; release(); } } };
}

function base14Covers(name: StandardPdfFont, text: string): boolean {
  if (name === 'Symbol' || name === 'ZapfDingbats') return false;
  for (const character of text) {
    const scalar = character.codePointAt(0)!;
    if (scalar < 0x20 || scalar > 0xff) return false;
  }
  return text.length > 0;
}

export function canDrawText(...args: any[]): boolean {
  if (typeof args[0] === 'string') return [...args[0]].every((character) => /[\x20-\x7e]/u.test(character));
  const [font, text = typeof args[0] === 'string' ? args[0] : ''] = args as [{ readonly chars?: Iterable<string | number>; readonly text?: string } | undefined, string?];
  if (!font) return false;
  if (typeof font.text === 'string') return [...text].every((character) => font.text!.includes(character));
  if (font.chars) {
    const supported = new Set(font.chars);
    return [...text].every((character) => supported.has(character) || supported.has(character.codePointAt(0)!));
  }
  return false;
}

export function listEditFonts<T extends { readonly font?: unknown }>(objects: readonly T[] = []): readonly unknown[] {
  if (objects.length === 0) return ['arial', 'times', 'courier'];
  const fonts: unknown[] = [];
  for (const object of objects) if (object.font !== undefined && !fonts.includes(object.font)) fonts.push(object.font);
  return fonts;
}
