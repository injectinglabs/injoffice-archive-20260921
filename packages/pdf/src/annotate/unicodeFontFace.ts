const maxFontBytes = 16 * 1024 * 1024
const maxFaces = 64
const maxTables = 4095
const align4 = (value: number) => Math.ceil(value / 4) * 4
interface Range { start: number; end: number }
interface Table extends Range { tag: number; length: number }
const overlaps = (a: Range, b: Range) => a.start < b.end && b.start < a.end

/** Big-endian padded uint32 checksum; head.checkSumAdjustment is zero in its table checksum. */
export function fontChecksum(bytes: Uint8Array, zeroAdjustment = false): number {
  let sum = 0
  for (let offset = 0; offset < bytes.length; offset += 4) {
    const word = zeroAdjustment && offset === 8 ? 0 : ((bytes[offset] ?? 0) * 0x1000000 + (bytes[offset + 1] ?? 0) * 0x10000 + (bytes[offset + 2] ?? 0) * 0x100 + (bytes[offset + 3] ?? 0))
    sum = (sum + word) >>> 0
  }
  return sum
}

/** Extracts a selected collection face without changing caller bytes or glyph IDs.
 * Standalone faces retain their original program bytes after directory validation. */
export function standaloneUnicodeFontFace(source: Uint8Array, faceIndex = 0): Uint8Array {
  if (!(source instanceof Uint8Array) || source.length < 12 || source.length > maxFontBytes) throw new RangeError('embedded appearance font must contain 12..16777216 bytes')
  if (!Number.isSafeInteger(faceIndex) || faceIndex < 0) throw new RangeError('font face index must be a nonnegative integer')
  const view = new DataView(source.buffer, source.byteOffset, source.byteLength)
  const ensure = (offset: number, length: number) => {
    if (!Number.isSafeInteger(offset) || !Number.isSafeInteger(length) || offset < 0 || length < 0 || offset + length > source.length) throw new RangeError('font table range exceeds supplied bytes')
  }
  const u16 = (offset: number) => { ensure(offset, 2); return view.getUint16(offset) }
  const u32 = (offset: number) => { ensure(offset, 4); return view.getUint32(offset) }
  const collection = u32(0) === 0x74746366
  const directories: Range[] = []
  let faceOffset = 0, signature: Range | undefined
  if (collection) {
    const version = u32(4), count = u32(8)
    if (![0x10000, 0x20000].includes(version) || count < 1 || count > maxFaces || faceIndex >= count) throw new RangeError('unsupported collection header or font face index')
    const headerLength = 12 + count * 4 + (version === 0x20000 ? 12 : 0)
    ensure(0, headerLength)
    directories.push({ start: 0, end: headerLength })
    for (let index = 0; index < count; index++) {
      const offset = u32(12 + index * 4), tableCount = u16(offset + 4)
      if (offset % 4 || ![0x10000, 0x74727565, 0x4f54544f].includes(u32(offset)) || tableCount < 1 || tableCount > maxTables) throw new RangeError('invalid collection face directory')
      ensure(offset, 12 + tableCount * 16)
      const range = { start: offset, end: offset + 12 + tableCount * 16 }
      if (directories.some(existing => overlaps(range, existing) && (range.start !== existing.start || range.end !== existing.end))) throw new RangeError('overlapping collection directories')
      directories.push(range)
      if (index === faceIndex) faceOffset = offset
    }
    if (version === 0x20000) {
      const position = 12 + count * 4, tag = u32(position), length = u32(position + 4), offset = u32(position + 8)
      if (tag || length || offset) {
        if (tag !== 0x44534947 || length < 8 || offset % 4) throw new RangeError('invalid collection signature range')
        ensure(offset, length); signature = { start: offset, end: offset + length }
        if (directories.some(directory => overlaps(signature!, directory))) throw new RangeError('collection signature overlaps a directory')
      }
    }
  } else if (faceIndex !== 0) throw new RangeError('standalone fonts only have face index 0')
  const scaler = u32(faceOffset), count = u16(faceOffset + 4)
  if (![0x10000, 0x74727565, 0x4f54544f].includes(scaler) || count < 1 || count > maxTables) throw new RangeError('invalid font table directory')
  ensure(faceOffset, 12 + count * 16)
  if (!collection) directories.push({ start: 0, end: 12 + count * 16 })
  const tags = new Set<number>(), tables: Table[] = []
  for (let index = 0; index < count; index++) {
    const record = faceOffset + 12 + index * 16, tag = u32(record), offset = u32(record + 8), length = u32(record + 12)
    const name = String.fromCharCode(...source.subarray(record, record + 4))
    if (!/^[!-~]{1,4} *$/.test(name)) throw new RangeError('invalid font table tag')
    if (tags.has(tag)) throw new RangeError('duplicate font table tag')
    tags.add(tag); ensure(offset, length)
    if (offset % 4) throw new RangeError('unaligned font table')
    const table = { tag, start: offset, end: offset + length, length }
    if (directories.some(directory => overlaps(table, directory) || offset >= directory.start && offset < directory.end) || signature && overlaps(table, signature)) throw new RangeError('font table overlaps container metadata')
    tables.push(table)
  }
  const sorted = [...tables].sort((a, b) => a.start - b.start)
  let previousEnd = 0
  for (const table of sorted) {
    if (table.start < previousEnd) throw new RangeError('overlapping font tables')
    previousEnd = Math.max(previousEnd, table.end)
  }
  const head = tables.find(table => table.tag === 0x68656164)
  if (!head || head.length < 54 || u32(head.start + 12) !== 0x5f0f3cf5) throw new RangeError('invalid font head table')
  if (!collection) return source.slice()
  // A copied face is a new SFNT file; the collection/face signatures no longer
  // describe it. Preserve the original caller collection and omit DSIG here.
  const selected = tables.filter(table => table.tag !== 0x44534947).sort((a, b) => a.tag - b.tag)
  const directoryLength = 12 + selected.length * 16
  const length = selected.reduce((size, table) => size + align4(table.length), directoryLength)
  if (length > maxFontBytes) throw new RangeError('extracted font exceeds embedded font byte limit')
  const output = new Uint8Array(length), out = new DataView(output.buffer)
  out.setUint32(0, scaler); out.setUint16(4, selected.length)
  const power = Math.floor(Math.log2(selected.length)), searchRange = 2 ** power * 16
  out.setUint16(6, searchRange); out.setUint16(8, power); out.setUint16(10, selected.length * 16 - searchRange)
  let offset = directoryLength, adjustment = 0
  selected.forEach((table, index) => {
    output.set(source.subarray(table.start, table.end), offset)
    const isHead = table.tag === 0x68656164
    if (isHead) { adjustment = offset + 8; out.setUint32(adjustment, 0) }
    const record = 12 + index * 16
    out.setUint32(record, table.tag); out.setUint32(record + 4, fontChecksum(output.subarray(offset, offset + table.length), isHead)); out.setUint32(record + 8, offset); out.setUint32(record + 12, table.length)
    offset += align4(table.length)
  })
  out.setUint32(adjustment, (0xb1b0afba - fontChecksum(output)) >>> 0)
  return output
}
