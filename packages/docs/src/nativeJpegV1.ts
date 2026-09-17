/** Bounded baseline JFIF marker validation (ITU-T T.81 / T.871).
 * No ICC profiles, Adobe transforms, progressive/multiscan, arithmetic coding
 * or external tables. One Exif attribute segment is read far enough to prove it
 * states no orientation but the identity; a rotating orientation is refused,
 * because a browser applies it and a rotated paint is not the source page.
 * The frame must be structurally complete through its EOI marker; what the
 * byte string carries past that marker is unreachable to any decoder and is
 * refused only when it would make the string two pictures instead of one.
 * Pixels remain the browser's job.
 */
/** True only for an APP1 payload that is a structurally complete Exif TIFF
 * header whose IFD0 either omits Orientation (TIFF 6.0 tag 0x0112) or states
 * the identity. Sub-IFDs are not walked: renderers read orientation from IFD0,
 * and IFD1 describes the thumbnail this module never paints. */
function identityOrientationExif(bytes: Uint8Array, start: number, end: number): boolean {
  if (end - start < 14 || String.fromCharCode(...bytes.subarray(start, start + 6)) !== 'Exif\0\0') return false
  const tiff = start + 6
  const big = bytes[tiff] === 0x4d && bytes[tiff + 1] === 0x4d
  if (!big && !(bytes[tiff] === 0x49 && bytes[tiff + 1] === 0x49)) return false
  const short = (at: number) => big ? bytes[at]! * 256 + bytes[at + 1]! : bytes[at + 1]! * 256 + bytes[at]!
  const long = (at: number) => big ? short(at) * 65536 + short(at + 2) : short(at + 2) * 65536 + short(at)
  if (short(tiff + 2) !== 42) return false
  const first = long(tiff + 4)
  if (first < 8 || tiff + first + 2 > end) return false
  const directory = tiff + first
  const count = short(directory)
  // TIFF stores an inline SHORT left-justified in its four value bytes, so the
  // orientation reads the same way under either byte order.
  if (count === 0 || count > 256 || directory + 2 + count * 12 + 4 > end) return false
  for (let index = 0; index < count; index += 1) {
    const entry = directory + 2 + index * 12
    if (short(entry) === 0x0112 && (short(entry + 2) !== 3 || long(entry + 4) !== 1 || short(entry + 8) !== 1)) return false
  }
  return true
}

export function nativeBaselineJpegDimensions(bytes: Uint8Array): { width: number; height: number } | undefined {
  if (bytes.length < 24 || bytes.length > 16 * 1024 * 1024 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return undefined
  let offset = 2, segments = 0, width = 0, height = 0, restartInterval = 0
  let jfif = false, exif = false
  const components = new Map<number, number>()
  const quantization = new Set<number>(), huffman = new Set<number>()
  const word = (at: number) => bytes[at]! * 256 + bytes[at + 1]!
  while (offset + 4 <= bytes.length && ++segments <= 256) {
    if (bytes[offset++] !== 0xff) return undefined
    const marker = bytes[offset++]!
    const length = word(offset)
    if (length < 2 || offset + length > bytes.length) return undefined
    const start = offset + 2, end = offset + length
    if (!jfif && marker !== 0xe0) return undefined
    if (marker === 0xe0) {
      // T.871 asks both densities to be nonzero, but density is display
      // metadata no decoder turns into pixels and this module never reads:
      // the picture is sized by its DrawingML extent. libgd's encoder has
      // long written the 0/0 pair, so a whole page is not worth that. A
      // half-zero pair still refuses: it states an aspect ratio and then
      // contradicts it.
      if (jfif || length !== 16 || String.fromCharCode(...bytes.subarray(start, start + 5)) !== 'JFIF\0' || bytes[start + 5] !== 1 || bytes[start + 6]! > 2 || bytes[start + 7]! > 2 || (word(start + 8) === 0) !== (word(start + 10) === 0) || bytes[start + 12] !== 0 || bytes[start + 13] !== 0) return undefined
      jfif = true
    } else if (marker === 0xe1) {
      // T.871 places Exif attributes in APP1. One segment, before the frame
      // header, and only when its IFD0 states no rotating orientation.
      if (exif || components.size || !identityOrientationExif(bytes, start, end)) return undefined
      exif = true
    } else if (marker === 0xc0) {
      if (components.size || bytes[start] !== 8) return undefined
      height = word(start + 1); width = word(start + 3)
      const count = bytes[start + 5]!
      if (!width || !height || ![1, 3].includes(count) || length !== 8 + 3 * count) return undefined
      let blocks = 0
      for (let i = 0; i < count; i++) {
        const at = start + 6 + i * 3, id = bytes[at]!, sampling = bytes[at + 1]!, table = bytes[at + 2]!
        const horizontal = sampling >>> 4, vertical = sampling & 15
        if (id !== i + 1 || horizontal < 1 || horizontal > 4 || vertical < 1 || vertical > 4 || table > 3) return undefined
        blocks += horizontal * vertical; components.set(id, table)
      }
      if (blocks > 10) return undefined
    } else if (marker === 0xdb) {
      for (let at = start; at < end;) {
        const table = bytes[at++]!
        if (table > 3 || at + 64 > end || quantization.has(table)) return undefined
        for (let i = 0; i < 64; i++) if (bytes[at + i] === 0) return undefined
        quantization.add(table); at += 64
      }
    } else if (marker === 0xc4) {
      for (let at = start; at < end;) {
        const table = bytes[at++]!
        if (table >>> 4 > 1 || (table & 15) > 3 || huffman.has(table) || at + 16 > end) return undefined
        let count = 0, available = 1
        // JPEG reserves an all-ones code at every length for entropy padding.
        // A complete tree (available === 0) is therefore invalid as well.
        for (let i = 0; i < 16; i++) { const n = bytes[at + i]!; count += n; available = available * 2 - n; if (available <= 0) return undefined }
        at += 16
        if (count === 0 || count > 256 || at + count > end) return undefined
        if (table >>> 4 === 0 && bytes.subarray(at, at + count).some((symbol) => symbol > 11)) return undefined
        if (table >>> 4 === 1 && bytes.subarray(at, at + count).some((symbol) => {
          const size = symbol & 15
          return size > 10 || (size === 0 && symbol !== 0x00 && symbol !== 0xf0)
        })) return undefined
        huffman.add(table); at += count
      }
    } else if (marker === 0xdd) {
      if (length !== 4) return undefined
      restartInterval = word(start)
    } else if (marker === 0xda) {
      const count = bytes[start]!
      if (!components.size || count !== components.size || length !== 6 + 2 * count) return undefined
      for (let i = 0; i < count; i++) {
        const id = bytes[start + 1 + i * 2]!, tables = bytes[start + 2 + i * 2]!
        if (id !== i + 1 || !quantization.has(components.get(id)!) || !huffman.has(tables >>> 4) || !huffman.has(0x10 | (tables & 15))) return undefined
      }
      if (bytes[end - 3] !== 0 || bytes[end - 2] !== 63 || bytes[end - 1] !== 0) return undefined
      let entropy = 0, nextRestart = 0
      for (let at = end; at < bytes.length;) {
        if (bytes[at++] !== 0xff) { entropy++; continue }
        if (at >= bytes.length) return undefined
        const code = bytes[at++]!
        if (code === 0) { entropy++; continue }
        if (code >= 0xd0 && code <= 0xd7) { if (!restartInterval || code !== 0xd0 + nextRestart) return undefined; nextRestart = (nextRestart + 1) % 8; continue }
        if (code !== 0xd9 || entropy === 0) return undefined
        // T.81 B.2: EOI terminates the compressed image data. Every decoder
        // stops there, and sniffing reads the SOI at offset 0, so bytes after
        // it are unreachable and cannot change the picture this byte string
        // is. A second SOI can: that string is two pictures, and choosing one
        // is not this module's call.
        for (let tail = at; tail + 1 < bytes.length; tail += 1) if (bytes[tail] === 0xff && bytes[tail + 1] === 0xd8) return undefined
        return { width, height }
      }
      return undefined
    } else if (marker !== 0xfe) return undefined
    offset = end
  }
  return undefined
}
