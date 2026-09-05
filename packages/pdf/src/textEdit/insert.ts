import { loadRequestedFont, type FontRequest, type StandardPdfFont } from './fontResolve.js';
import { PdfiumSession } from './pdfium.js';
import { createStyledTextObject } from './rebuild.js';
import { assertEditableText } from './splice.js';

const MAX_OPERATIONS = 1_000;

export interface TextInsertSpec extends FontRequest {
  readonly page: number;
  readonly text: string;
  readonly x?: number;
  readonly y?: number;
  readonly origin?: readonly [number, number];
  readonly fontSize: number;
  readonly color?: readonly [number, number, number] | readonly [number, number, number, number];
  readonly fontName?: StandardPdfFont;
  readonly font?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly rotate?: number;
  readonly lineLeading?: number;
}

export interface TextInsertFailure {
  readonly index: number;
  readonly page?: number;
  readonly code: 'invalid-spec' | 'unsupported-font' | 'pdfium-error';
  readonly message: string;
  readonly reason?: string;
}

export interface TextInsertsResult {
  readonly pdfBytes: Uint8Array;
  readonly bytes: Uint8Array;
  readonly pdf: Uint8Array;
  readonly applied: number;
  readonly skipped: readonly TextInsertFailure[];
  readonly failures: readonly TextInsertFailure[];
}

export async function applyTextInserts(pdfBytes: Uint8Array, inserts: readonly TextInsertSpec[]): Promise<TextInsertsResult> {
  const original = cloneInput(pdfBytes);
  if (!Array.isArray(inserts) || inserts.length > MAX_OPERATIONS) {
    return result(original, 0, [{ index: 0, code: 'invalid-spec', message: `inserts must be an array of at most ${MAX_OPERATIONS} items` }]);
  }
  const validation = inserts.flatMap((insert, index) => validateInsert(insert, index));
  if (validation.length) return result(original, 0, validation);
  if (inserts.length === 0) return result(original, 0, []);

  const session = await PdfiumSession.open(original);
  try {
    for (let index = 0; index < inserts.length; index += 1) {
      const insert = inserts[index]!;
      try { applyOneInsert(session, insert); }
      catch (error) { return result(original, 0, [{ index, page: insert.page, code: classifyFontFailure(error), message: errorMessage(error) }]); }
    }
    return result(session.save(), inserts.length, []);
  } finally {
    session.close();
  }
}

function applyOneInsert(session: PdfiumSession, insert: TextInsertSpec): void {
  session.withPage(insert.page, (page) => {
    const lines = insert.text.split(/\r?\n/u);
    const font = loadRequestedFont(session, lines.join(''), {
      fontBytes: insert.fontBytes,
      fontFaceIndex: insert.fontFaceIndex,
      standardFont: styledStandardFont(insert.standardFont ?? insert.fontName ?? insert.font, insert.bold, insert.italic),
    });
    try {
      const fill = normalizeColor(insert.color);
      const angle = (insert.rotate ?? 0) * Math.PI / 180;
      const leading = insert.lineLeading ?? insert.fontSize * 1.2;
      lines.forEach((line, index) => {
        if (!line) return;
        const object = createStyledTextObject(session, line, font, {
          font: font.handle,
          fontSize: insert.fontSize,
          fill,
          renderMode: 0,
          matrix: { a: Math.cos(angle), b: Math.sin(angle), c: -Math.sin(angle), d: Math.cos(angle), e: insert.x ?? insert.origin![0], f: (insert.y ?? insert.origin![1]) - index * leading },
        });
        session.api.FPDFPage_InsertObject(page, object);
      });
      session.generate(page);
    } finally {
      font.close();
    }
  });
}

function validateInsert(insert: TextInsertSpec, index: number): TextInsertFailure[] {
  try {
    if (!insert || typeof insert !== 'object') throw new TypeError('insert must be an object');
    assertEditableText(insert.text, 'text');
    if (insert.text.trim().length === 0) throw new RangeError('empty inserted text is not supported');
    if (!Number.isSafeInteger(insert.page) || insert.page < 1) throw new RangeError('page must be a positive integer');
    const x = insert.x ?? insert.origin?.[0];
    const y = insert.y ?? insert.origin?.[1];
    for (const [name, value] of [['x', x], ['y', y], ['fontSize', insert.fontSize]] as const) {
      if (typeof value !== 'number' || !Number.isFinite(value) || Math.abs(value) > 10_000_000 || (name === 'fontSize' && value <= 0)) throw new RangeError(`${name} is outside the supported range`);
    }
    normalizeColor(insert.color);
    if (insert.rotate !== undefined && !Number.isFinite(insert.rotate)) throw new RangeError('rotate must be finite degrees');
    if (insert.lineLeading !== undefined && (!Number.isFinite(insert.lineLeading) || insert.lineLeading <= 0)) throw new RangeError('lineLeading must be positive');
    if (insert.fontBytes && insert.standardFont) throw new RangeError('fontBytes and standardFont are mutually exclusive');
    return [];
  } catch (error) {
    return [{ index, page: insert?.page, code: 'invalid-spec', message: errorMessage(error) }];
  }
}

function normalizeColor(color: TextInsertSpec['color']): readonly [number, number, number, number] {
  const values = color ?? [0, 0, 0, 255];
  if (values.length !== 3 && values.length !== 4) throw new RangeError('color must have three or four channels');
  const rgba = [values[0], values[1], values[2], values.length === 4 ? values[3] : 255];
  if (!rgba.every((value) => Number.isInteger(value) && value! >= 0 && value! <= 255)) throw new RangeError('color channels must be integers from 0 to 255');
  return rgba as [number, number, number, number];
}

function styledStandardFont(name: string | undefined, bold = false, italic = false): StandardPdfFont | undefined {
  if (!name) return undefined;
  name = name.toLowerCase() === 'arial' ? 'Helvetica' : name.toLowerCase() === 'times' ? 'Times-Roman' : name.toLowerCase() === 'courier' ? 'Courier' : name;
  if (name === 'Helvetica' || name === 'Courier') {
    if (bold && italic) return `${name}-BoldOblique` as StandardPdfFont;
    if (bold) return `${name}-Bold` as StandardPdfFont;
    if (italic) return `${name}-Oblique` as StandardPdfFont;
  }
  if (name === 'Times-Roman') {
    if (bold && italic) return 'Times-BoldItalic';
    if (bold) return 'Times-Bold';
    if (italic) return 'Times-Italic';
  }
  return name as StandardPdfFont;
}

function cloneInput(bytes: Uint8Array): Uint8Array {
  if (!(bytes instanceof Uint8Array)) throw new TypeError('pdfBytes must be a Uint8Array');
  return bytes;
}

function result(pdfBytes: Uint8Array, applied: number, failures: readonly TextInsertFailure[]): TextInsertsResult {
  const skipped = failures.map((failure) => ({ ...failure, reason: failure.message }));
  return { pdfBytes, bytes: pdfBytes, pdf: pdfBytes, applied, skipped, failures };
}

function classifyFontFailure(error: unknown): TextInsertFailure['code'] {
  return error instanceof RangeError && /font|glyph|ASCII|sfnt|TrueType/i.test(error.message) ? 'unsupported-font' : 'pdfium-error';
}

function errorMessage(error: unknown): string { return error instanceof Error ? error.message : String(error); }
