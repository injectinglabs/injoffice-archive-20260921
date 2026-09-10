/** Bounded baseline JFIF marker validation (ITU-T T.81 / T.871).
 * No EXIF orientation, ICC profiles, Adobe transforms, progressive/multiscan,
 * arithmetic coding or external tables. Pixels remain the browser's job.
 */
export function nativeBaselineJpegDimensions(bytes: Uint8Array): { width: number; height: number } | undefined {
  if (bytes.length < 24 || bytes.length > 16 * 1024 * 1024 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return undefined
  let offset = 2, segments = 0, width = 0, height = 0, restartInterval = 0
  let jfif = false
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
      if (jfif || length !== 16 || String.fromCharCode(...bytes.subarray(start, start + 5)) !== 'JFIF\0' || bytes[start + 5] !== 1 || bytes[start + 6]! > 2 || bytes[start + 7]! > 2 || word(start + 8) === 0 || word(start + 10) === 0 || bytes[start + 12] !== 0 || bytes[start + 13] !== 0) return undefined
      jfif = true
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
        return code === 0xd9 && entropy > 0 && at === bytes.length ? { width, height } : undefined
      }
      return undefined
    } else if (marker !== 0xfe) return undefined
    offset = end
  }
  return undefined
}
