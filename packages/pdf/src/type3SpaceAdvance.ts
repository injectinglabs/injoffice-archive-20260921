import { PDFArray, PDFDict, PDFDocument, PDFName, PDFNumber, PDFRef, PDFStream } from 'pdf-lib';

const SPACE = 32;
const SPACE_NAME = 'space';
const NAMED_ENCODINGS_WITH_SPACE = new Set(['StandardEncoding', 'WinAnsiEncoding', 'MacRomanEncoding', 'MacExpertEncoding']);

export type Type3SpaceDecision =
  | { action: 'advance'; width: number; glyphName: typeof SPACE_NAME }
  | { action: 'omit'; reason: string };

function nameOf(value: unknown): string | undefined {
  if (!(value instanceof PDFName)) return undefined;
  return value.asString().replace(/^\//, '');
}

function numberOf(value: unknown): number | undefined {
  if (!(value instanceof PDFNumber)) return undefined;
  const amount = value.asNumber();
  return Number.isFinite(amount) ? amount : undefined;
}

function formatPdfNumber(value: number): string {
  if (Number.isInteger(value) && Math.abs(value) <= Number.MAX_SAFE_INTEGER) return String(value);
  return value.toFixed(6).replace(/\.?0+$/, '');
}

function containsAscii(bytes: Uint8Array, ascii: string): boolean {
  const first = ascii.charCodeAt(0);
  const last = bytes.length - ascii.length;
  outer: for (let i = 0; i <= last; i++) {
    if (bytes[i] !== first) continue;
    for (let j = 1; j < ascii.length; j++) {
      if (bytes[i + j] !== ascii.charCodeAt(j)) continue outer;
    }
    return true;
  }
  return false;
}

function glyphNameAt32(encoding: unknown): string | undefined {
  if (encoding instanceof PDFName) {
    const named = nameOf(encoding);
    return named && NAMED_ENCODINGS_WITH_SPACE.has(named) ? SPACE_NAME : undefined;
  }
  if (!(encoding instanceof PDFDict)) return undefined;
  const table = new Array<string | undefined>(256);
  const base = nameOf(encoding.lookupMaybe(PDFName.of('BaseEncoding'), PDFName));
  if (!base || NAMED_ENCODINGS_WITH_SPACE.has(base)) table[SPACE] = SPACE_NAME;
  const differences = encoding.lookupMaybe(PDFName.of('Differences'), PDFArray);
  if (!differences) return table[SPACE];
  let code = 0;
  for (let i = 0; i < differences.size(); i++) {
    const item = differences.lookup(i);
    if (item instanceof PDFNumber) {
      const next = item.asNumber();
      if (!Number.isInteger(next) || next < 0 || next > 255) return undefined;
      code = next;
      continue;
    }
    const glyph = nameOf(item);
    if (!glyph || code > 255) return undefined;
    table[code] = glyph;
    code += 1;
  }
  return table[SPACE];
}

function widthFor32(font: PDFDict): number | undefined {
  const firstChar = numberOf(font.lookup(PDFName.of('FirstChar')));
  const lastChar = numberOf(font.lookup(PDFName.of('LastChar')));
  const widths = font.lookupMaybe(PDFName.of('Widths'), PDFArray);
  if (
    widths &&
    firstChar !== undefined &&
    lastChar !== undefined &&
    Number.isInteger(firstChar) &&
    Number.isInteger(lastChar) &&
    firstChar <= SPACE &&
    SPACE <= lastChar
  ) {
    const index = SPACE - firstChar;
    if (index < widths.size()) {
      const width = numberOf(widths.lookup(index));
      if (width !== undefined) return width;
    }
  }
  const descriptor = font.lookup(PDFName.of('FontDescriptor'));
  if (!(descriptor instanceof PDFDict)) return undefined;
  return numberOf(descriptor.lookup(PDFName.of('MissingWidth')));
}

export function decideType3SpaceAdvance(font: PDFDict): Type3SpaceDecision {
  const subtype = nameOf(font.lookupMaybe(PDFName.of('Subtype'), PDFName));
  if (subtype !== 'Type3') return { action: 'omit', reason: 'not a Type3 font' };
  const type = nameOf(font.lookupMaybe(PDFName.of('Type'), PDFName));
  if (type !== undefined && type !== 'Font') return { action: 'omit', reason: 'not a font dictionary' };
  if (glyphNameAt32(font.lookup(PDFName.of('Encoding'))) !== SPACE_NAME) {
    return { action: 'omit', reason: 'encoding does not map character 32 to /space' };
  }
  const width = widthFor32(font);
  if (width === undefined) return { action: 'omit', reason: 'no Widths or MissingWidth for space' };
  return { action: 'advance', width, glyphName: SPACE_NAME };
}

function charProcsOf(font: PDFDict): PDFDict | undefined {
  const charProcs = font.lookup(PDFName.of('CharProcs'));
  return charProcs instanceof PDFDict ? charProcs : undefined;
}

function injectAdvanceOnlySpace(doc: PDFDocument, font: PDFDict, width: number): boolean {
  const charProcs = charProcsOf(font);
  if (!charProcs || charProcs.has(PDFName.of(SPACE_NAME))) return false;
  const stream = doc.context.stream(`${formatPdfNumber(width)} 0 d0\n`);
  charProcs.set(PDFName.of(SPACE_NAME), doc.context.register(stream));
  return true;
}

function shouldInspect(bytes: Uint8Array): boolean {
  return containsAscii(bytes, 'Type3') || containsAscii(bytes, 'ObjStm');
}

/** Inserts a width-only Type3 /space CharProc when the font proves that width.
 * Existing CharProcs and fonts without a proven space width are left unchanged. */
export async function ensureType3SpaceAdvances(bytes: Uint8Array): Promise<Uint8Array> {
  if (bytes.byteLength === 0 || !shouldInspect(bytes)) return bytes;
  let doc: PDFDocument;
  try {
    doc = await PDFDocument.load(bytes, { updateMetadata: false });
  } catch {
    return bytes;
  }
  if (doc.isEncrypted) return bytes;
  let changed = false;
  for (const [, object] of doc.context.enumerateIndirectObjects()) {
    if (!(object instanceof PDFDict)) continue;
    const decision = decideType3SpaceAdvance(object);
    if (decision.action !== 'advance') continue;
    if (injectAdvanceOnlySpace(doc, object, decision.width)) changed = true;
  }
  return changed ? new Uint8Array(await doc.save({ updateFieldAppearances: false })) : bytes;
}

export function type3SpaceCharProc(doc: PDFDocument, font: PDFDict): PDFStream | undefined {
  const charProcs = charProcsOf(font);
  const value = charProcs?.get(PDFName.of(SPACE_NAME));
  if (value instanceof PDFStream) return value;
  if (value instanceof PDFRef) {
    const stream = doc.context.lookup(value);
    return stream instanceof PDFStream ? stream : undefined;
  }
  return undefined;
}
