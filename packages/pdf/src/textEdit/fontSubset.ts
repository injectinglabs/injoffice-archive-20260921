import { fontCoversText, fontHasTrueTypeOutlines, readFontFace } from './fontCmap.js';
import subsetFont from 'subset-font';

export const MAX_FONT_BYTES = 64 * 1024 * 1024;

export interface PreparedFont {
  readonly bytes: Uint8Array;
  readonly faceIndex: number;
  readonly subset: false;
}

export function prepareFont(fontBytes: Uint8Array, text: string, faceIndex = 0): PreparedFont {
  if (!(fontBytes instanceof Uint8Array)) throw new TypeError('fontBytes must be a Uint8Array');
  if (fontBytes.byteLength === 0 || fontBytes.byteLength > MAX_FONT_BYTES) throw new RangeError(`fontBytes must contain 1..${MAX_FONT_BYTES} bytes`);
  readFontFace(fontBytes, faceIndex);
  if (!fontHasTrueTypeOutlines(fontBytes, faceIndex)) throw new RangeError('v1 accepts only fonts with TrueType glyf outlines');
  if (!fontCoversText(fontBytes, text, faceIndex)) throw new RangeError('font does not map every requested Unicode scalar');
  return { bytes: fontBytes.slice(), faceIndex, subset: false };
}

export async function subsetTtf(fontBytes: Uint8Array, text = '', faceIndex = 0): Promise<Buffer> {
  if (!(fontBytes instanceof Uint8Array)) throw new TypeError('fontBytes must be a Uint8Array');
  if (fontBytes.byteLength === 0 || fontBytes.byteLength > MAX_FONT_BYTES) throw new RangeError(`fontBytes must contain 1..${MAX_FONT_BYTES} bytes`);
  const face = readFontFace(fontBytes, faceIndex);
  if (face.offset !== 0) throw new RangeError('collection subsetting is not supported by the documented subset-font API');
  if (text && !fontCoversText(fontBytes, text, faceIndex)) throw new RangeError('font does not map every requested Unicode scalar');
  return subsetFont(Buffer.from(fontBytes), text, {
    targetFormat: 'sfnt',
    preserveNameIds: [0, 1, 2, 3, 4, 5, 6],
  });
}

export function identityCffCharset(input: number | Uint8Array): Buffer {
  if (input instanceof Uint8Array) return Buffer.isBuffer(input) ? input : Buffer.from(input);
  const glyphCount = input;
  if (!Number.isSafeInteger(glyphCount) || glyphCount < 1 || glyphCount > 65_535) throw new RangeError('glyphCount is outside the supported CFF charset range');
  const bytes = new Uint8Array(1 + Math.max(0, glyphCount - 1) * 2);
  bytes[0] = 0;
  for (let glyph = 1; glyph < glyphCount; glyph += 1) {
    const sid = glyph - 1;
    bytes[1 + (glyph - 1) * 2] = sid >>> 8;
    bytes[2 + (glyph - 1) * 2] = sid & 0xff;
  }
  return Buffer.from(bytes);
}
