const MAX_TABLES = 4096;

export interface FontFaceView {
  readonly offset: number;
  readonly tables: ReadonlyMap<string, { offset: number; length: number }>;
}

class BinaryView {
  readonly data: Uint8Array;
  readonly view: DataView;

  constructor(data: Uint8Array) {
    this.data = data;
    this.view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  }

  ensure(offset: number, size: number): void {
    if (!Number.isSafeInteger(offset) || !Number.isSafeInteger(size) || offset < 0 || size < 0 || offset + size > this.data.length) {
      throw new RangeError('font table offset is outside the supplied bytes');
    }
  }

  u16(offset: number): number { this.ensure(offset, 2); return this.view.getUint16(offset, false); }
  i16(offset: number): number { this.ensure(offset, 2); return this.view.getInt16(offset, false); }
  u32(offset: number): number { this.ensure(offset, 4); return this.view.getUint32(offset, false); }
  tag(offset: number): string {
    this.ensure(offset, 4);
    return String.fromCharCode(this.data[offset]!, this.data[offset + 1]!, this.data[offset + 2]!, this.data[offset + 3]!);
  }
}

export function readFontFace(bytes: Uint8Array, faceIndex = 0): FontFaceView {
  if (!(bytes instanceof Uint8Array) || bytes.byteLength < 12) throw new RangeError('font bytes are too short');
  const binary = new BinaryView(bytes);
  let faceOffset = 0;
  if (binary.tag(0) === 'ttcf') {
    binary.ensure(0, 12);
    const count = binary.u32(8);
    if (!Number.isSafeInteger(faceIndex) || faceIndex < 0 || faceIndex >= count) throw new RangeError('font collection face index is out of range');
    binary.ensure(12, count * 4);
    faceOffset = binary.u32(12 + faceIndex * 4);
  } else if (faceIndex !== 0) {
    throw new RangeError('non-collection fonts only have face index 0');
  }
  binary.ensure(faceOffset, 12);
  const scaler = binary.u32(faceOffset);
  if (scaler !== 0x00010000 && scaler !== 0x4f54544f && scaler !== 0x74727565) throw new RangeError('unsupported sfnt scaler type');
  const count = binary.u16(faceOffset + 4);
  if (count === 0 || count > MAX_TABLES) throw new RangeError('invalid sfnt table count');
  binary.ensure(faceOffset + 12, count * 16);
  const tables = new Map<string, { offset: number; length: number }>();
  for (let index = 0; index < count; index += 1) {
    const record = faceOffset + 12 + index * 16;
    const tag = binary.tag(record);
    const offset = binary.u32(record + 8);
    const length = binary.u32(record + 12);
    binary.ensure(offset, length);
    if (!tables.has(tag)) tables.set(tag, { offset, length });
  }
  return { offset: faceOffset, tables };
}

export function fontHasTrueTypeOutlines(bytes: Uint8Array, faceIndex = 0): boolean {
  try { return readFontFace(bytes, faceIndex).tables.has('glyf'); } catch { return false; }
}

export function glyphForScalar(bytes: Uint8Array, scalar: number, faceIndex = 0): number {
  if (!Number.isInteger(scalar) || scalar < 0 || scalar > 0x10ffff || (scalar >= 0xd800 && scalar <= 0xdfff)) return 0;
  const binary = new BinaryView(bytes);
  const cmap = readFontFace(bytes, faceIndex).tables.get('cmap');
  if (!cmap) return 0;
  binary.ensure(cmap.offset, Math.min(cmap.length, 4));
  const count = binary.u16(cmap.offset + 2);
  binary.ensure(cmap.offset + 4, count * 8);
  const choices: Array<{ priority: number; offset: number; format: number }> = [];
  for (let index = 0; index < count; index += 1) {
    const record = cmap.offset + 4 + index * 8;
    const platform = binary.u16(record);
    const encoding = binary.u16(record + 2);
    const relative = binary.u32(record + 4);
    if (relative >= cmap.length) continue;
    const offset = cmap.offset + relative;
    const format = binary.u16(offset);
    if (format !== 4 && format !== 12) continue;
    const priority = format === 12 && platform === 3 && encoding === 10 ? 0
      : format === 12 && platform === 0 ? 1
      : format === 4 && platform === 3 && encoding === 1 ? 2
      : format === 4 && platform === 0 ? 3 : 4;
    choices.push({ priority, offset, format });
  }
  choices.sort((left, right) => left.priority - right.priority);
  for (const choice of choices) {
    const glyph = choice.format === 12 ? lookupFormat12(binary, choice.offset, cmap, scalar) : lookupFormat4(binary, choice.offset, cmap, scalar);
    if (glyph !== 0) return glyph;
  }
  return 0;
}

export function fontCoversText(bytes: Uint8Array, text: string, faceIndex = 0): boolean {
  for (const character of text) {
    const scalar = character.codePointAt(0)!;
    if (scalar === 0x0a || scalar === 0x0d || scalar === 0x09) continue;
    if (glyphForScalar(bytes, scalar, faceIndex) === 0) return false;
  }
  return true;
}

function lookupFormat12(binary: BinaryView, offset: number, table: { offset: number; length: number }, scalar: number): number {
  binary.ensure(offset, 16);
  const length = binary.u32(offset + 4);
  const groups = binary.u32(offset + 12);
  if (length < 16 || offset + length > table.offset + table.length || groups > Math.floor((length - 16) / 12)) return 0;
  let low = 0;
  let high = groups - 1;
  while (low <= high) {
    const middle = Math.floor((low + high) / 2);
    const record = offset + 16 + middle * 12;
    const start = binary.u32(record);
    const end = binary.u32(record + 4);
    if (scalar < start) high = middle - 1;
    else if (scalar > end) low = middle + 1;
    else return (binary.u32(record + 8) + scalar - start) >>> 0;
  }
  return 0;
}

function lookupFormat4(binary: BinaryView, offset: number, table: { offset: number; length: number }, scalar: number): number {
  if (scalar > 0xffff) return 0;
  binary.ensure(offset, 14);
  const length = binary.u16(offset + 2);
  if (length < 16 || offset + length > table.offset + table.length) return 0;
  const segments = binary.u16(offset + 6) / 2;
  if (!Number.isInteger(segments) || segments < 1) return 0;
  const endCodes = offset + 14;
  const startCodes = endCodes + segments * 2 + 2;
  const deltas = startCodes + segments * 2;
  const rangeOffsets = deltas + segments * 2;
  binary.ensure(rangeOffsets, segments * 2);
  for (let index = 0; index < segments; index += 1) {
    const end = binary.u16(endCodes + index * 2);
    if (scalar > end) continue;
    const start = binary.u16(startCodes + index * 2);
    if (scalar < start) return 0;
    const delta = binary.i16(deltas + index * 2);
    const rangeOffset = binary.u16(rangeOffsets + index * 2);
    if (rangeOffset === 0) return (scalar + delta) & 0xffff;
    const glyphAddress = rangeOffsets + index * 2 + rangeOffset + (scalar - start) * 2;
    if (glyphAddress + 2 > offset + length) return 0;
    const glyph = binary.u16(glyphAddress);
    return glyph === 0 ? 0 : (glyph + delta) & 0xffff;
  }
  return 0;
}
