import type { PdfiumSession } from './pdfium.js';
import type { ResolvedFont } from './fontResolve.js';
import type { TextObjectStyle } from './styleRuns.js';

export interface WrappedLine {
  readonly text: string;
  readonly estimatedWidth: number;
}

export interface StyledSegment {
  readonly text: string;
  readonly font: ResolvedFont;
  readonly style: TextObjectStyle;
}

export function wrapText(text: string, maxWidth: number, fontSize: number): readonly WrappedLine[] {
  if (!Number.isFinite(maxWidth) || maxWidth <= 0) throw new RangeError('maxWidth must be a positive finite number');
  if (!Number.isFinite(fontSize) || fontSize <= 0) throw new RangeError('fontSize must be a positive finite number');
  const lines: WrappedLine[] = [];
  for (const paragraph of text.split('\n')) {
    if (paragraph.length === 0) { lines.push({ text: '', estimatedWidth: 0 }); continue; }
    let current = '';
    let width = 0;
    for (const token of paragraph.match(/\s+|\S+/gu) ?? []) {
      const tokenWidth = estimateTextWidth(token, fontSize);
      if (current && width + tokenWidth > maxWidth) {
        lines.push({ text: current.trimEnd(), estimatedWidth: estimateTextWidth(current.trimEnd(), fontSize) });
        current = token.trimStart();
        width = estimateTextWidth(current, fontSize);
      } else {
        current += token;
        width += tokenWidth;
      }
      while (width > maxWidth && [...current].length > 1) {
        const split = fittingPrefix(current, maxWidth, fontSize);
        lines.push({ text: split.head, estimatedWidth: estimateTextWidth(split.head, fontSize) });
        current = split.tail;
        width = estimateTextWidth(current, fontSize);
      }
    }
    lines.push({ text: current.trimEnd(), estimatedWidth: estimateTextWidth(current.trimEnd(), fontSize) });
  }
  return lines;
}

export function createStyledTextObject(
  session: PdfiumSession,
  text: string,
  font: ResolvedFont,
  style: TextObjectStyle,
  lineOffset = 0,
  horizontalOffset = 0,
): number {
  const object = session.api.FPDFPageObj_CreateTextObj(session.document, font.handle, style.fontSize);
  if (!object) throw new Error('PDFium could not create a text object');
  try {
    if (!session.assignText(object, text)) throw new Error('PDFium could not encode replacement text with the selected font');
    session.setTextObjectStyle(object, { ...style, matrix: { ...style.matrix, e: style.matrix.e + horizontalOffset } }, lineOffset);
    return object;
  } catch (error) {
    session.api.FPDFPageObj_Destroy(object);
    throw error;
  }
}

export function replaceObjectWithLines(
  session: PdfiumSession,
  page: number,
  sourceObject: number,
  lines: readonly WrappedLine[],
  font: ResolvedFont,
  style: TextObjectStyle,
  lineHeight: number,
  alignment: 'left' | 'center' | 'right',
  maxWidth: number,
  lineXOffsets: readonly number[] = [],
): void {
  const insertionIndex = session.findObjectIndex(page, sourceObject);
  if (insertionIndex < 0) throw new Error('source text object is not a direct page object');
  const created: number[] = [];
  try {
    for (let index = 0; index < lines.length; index += 1) {
      const line = lines[index]!;
      if (line.text.length === 0) continue;
      const alignedOffset = alignment === 'left' ? 0 : alignment === 'center' ? (maxWidth - line.estimatedWidth) / 2 : maxWidth - line.estimatedWidth;
      const horizontalOffset = alignedOffset + (lineXOffsets[index] ?? 0);
      created.push(createStyledTextObject(session, line.text, font, style, index * lineHeight, horizontalOffset));
    }
    if (!session.api.FPDFPage_RemoveObject(page, sourceObject)) throw new Error('PDFium could not remove the source text object');
    session.api.FPDFPageObj_Destroy(sourceObject);
    for (let index = 0; index < created.length; index += 1) {
      const object = created[index]!;
      if (!session.api.FPDFPage_InsertObjectAtIndex(page, object, insertionIndex + index)) {
        throw new Error('PDFium could not insert a rebuilt text object');
      }
      created[index] = 0;
    }
    created.length = 0;
  } finally {
    for (const object of created) if (object) session.api.FPDFPageObj_Destroy(object);
  }
}

export function replaceObjectsWithSegments(
  session: PdfiumSession,
  page: number,
  sourceObjects: readonly number[],
  segments: readonly StyledSegment[],
): void {
  if (sourceObjects.length === 0) throw new Error('segmented rebuild requires at least one source object');
  const sourceIndices = sourceObjects.map((object) => session.findObjectIndex(page, object));
  if (sourceIndices.some((index) => index < 0)) throw new Error('source text object is not a direct page object');
  const insertionIndex = Math.min(...sourceIndices);
  const created: number[] = [];
  let horizontalOffset = 0;
  try {
    for (const segment of segments) {
      if (!segment.text) continue;
      const object = createStyledTextObject(session, segment.text, segment.font, segment.style, 0, horizontalOffset);
      created.push(object);
      horizontalOffset += measureAdvance(session, segment, object);
    }
    for (const object of sourceObjects) {
      if (!session.api.FPDFPage_RemoveObject(page, object)) throw new Error('PDFium could not remove a source text object');
      session.api.FPDFPageObj_Destroy(object);
    }
    for (let index = 0; index < created.length; index += 1) {
      const object = created[index]!;
      if (!session.api.FPDFPage_InsertObjectAtIndex(page, object, insertionIndex + index)) throw new Error('PDFium could not insert a rebuilt text segment');
      created[index] = 0;
    }
    created.length = 0;
  } finally {
    for (const object of created) if (object) session.api.FPDFPageObj_Destroy(object);
  }
}

function measureAdvance(session: PdfiumSession, segment: StyledSegment, object: number): number {
  const measured = objectWidth(session, object);
  let combined = 0;
  let marker = 0;
  try {
    combined = createStyledTextObject(session, `${segment.text}M`, segment.font, segment.style);
    marker = createStyledTextObject(session, 'M', segment.font, segment.style);
    const advance = objectWidth(session, combined) - objectWidth(session, marker);
    return Number.isFinite(advance) && advance > 0 ? advance : measured;
  } catch {
    return measured > 0 ? measured : estimateTextWidth(segment.text, segment.style.fontSize);
  } finally {
    if (combined) session.api.FPDFPageObj_Destroy(combined);
    if (marker) session.api.FPDFPageObj_Destroy(marker);
  }
}

function objectWidth(session: PdfiumSession, object: number): number {
  const [left, , right] = session.readObjectBounds(object);
  const width = right - left;
  return Number.isFinite(width) && width > 0 ? width : 0;
}

function estimateTextWidth(text: string, fontSize: number): number {
  let units = 0;
  for (const character of text) {
    const scalar = character.codePointAt(0)!;
    units += scalar === 0x20 || scalar === 0x09 ? 0.32 : isWide(scalar) ? 1 : 0.56;
  }
  return units * fontSize;
}

function fittingPrefix(text: string, maxWidth: number, fontSize: number): { head: string; tail: string } {
  let head = '';
  for (const character of text) {
    if (head && estimateTextWidth(head + character, fontSize) > maxWidth) break;
    head += character;
  }
  return { head, tail: text.slice(head.length) };
}

function isWide(scalar: number): boolean {
  return (scalar >= 0x1100 && scalar <= 0x115f)
    || (scalar >= 0x2e80 && scalar <= 0xa4cf)
    || (scalar >= 0xac00 && scalar <= 0xd7a3)
    || (scalar >= 0xf900 && scalar <= 0xfaff)
    || (scalar >= 0x1f300 && scalar <= 0x1faff);
}
