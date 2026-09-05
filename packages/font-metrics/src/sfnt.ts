import { closeSync, existsSync, fstatSync, openSync, readFileSync, readSync, readdirSync, statSync } from 'node:fs'
import { homedir, platform } from 'node:os'
import { extname, join } from 'node:path'

export interface FaceNames {
  ps: string[]
  families: string[]
  subfamilies: string[]
}

export interface FaceRef {
  path: string
  offset: number
  style: string
}

export interface FontIndex {
  byPs: Map<string, FaceRef>
  byFamily: Map<string, FaceRef[]>
}

interface TableRecord {
  tag: string
  checksum: number
  offset: number
  length: number
}

interface ParsedFace {
  ref: FaceRef
  names: FaceNames
}

const TTC_TAG = 0x74746366
const OTTO_TAG = 0x4f54544f
const TRUE_TAG = 0x74727565
const SFNT_1 = 0x00010000
const MAX_FACES = 4096
const MAX_TABLES = 4096
const MAX_FONT_BYTES = 256 * 1024 * 1024
const FONT_EXTENSIONS = new Set(['.ttf', '.otf', '.ttc', '.otc'])

let cachedIndex: FontIndex | undefined

function contains(buffer: Buffer, offset: number, length: number): boolean {
  return Number.isSafeInteger(offset) && Number.isSafeInteger(length) && offset >= 0 && length >= 0 && offset <= buffer.length - length
}

function u16(buffer: Buffer, offset: number): number {
  if (!contains(buffer, offset, 2)) throw new RangeError('truncated sfnt uint16')
  return buffer.readUInt16BE(offset)
}

function u32(buffer: Buffer, offset: number): number {
  if (!contains(buffer, offset, 4)) throw new RangeError('truncated sfnt uint32')
  return buffer.readUInt32BE(offset)
}

function tagAt(buffer: Buffer, offset: number): string {
  if (!contains(buffer, offset, 4)) throw new RangeError('truncated sfnt tag')
  return buffer.toString('latin1', offset, offset + 4)
}

function isSfntFlavor(value: number): boolean {
  return value === SFNT_1 || value === OTTO_TAG || value === TRUE_TAG
}

function faceOffsets(buffer: Buffer): number[] {
  const signature = u32(buffer, 0)
  if (signature !== TTC_TAG) {
    if (!isSfntFlavor(signature)) throw new TypeError('unsupported sfnt signature')
    return [0]
  }

  const major = u16(buffer, 4)
  const minor = u16(buffer, 6)
  if ((major !== 1 && major !== 2) || minor !== 0) throw new TypeError('unsupported TTC version')
  const count = u32(buffer, 8)
  if (count === 0 || count > MAX_FACES || !contains(buffer, 12, count * 4)) throw new RangeError('invalid TTC face count')
  const offsets: number[] = []
  for (let i = 0; i < count; i++) {
    const offset = u32(buffer, 12 + i * 4)
    if (!contains(buffer, offset, 12) || !isSfntFlavor(u32(buffer, offset))) throw new RangeError('invalid TTC face offset')
    offsets.push(offset)
  }
  return offsets
}

function tableDirectory(buffer: Buffer, faceOffset: number): { flavor: number; records: TableRecord[] } {
  if (!contains(buffer, faceOffset, 12)) throw new RangeError('truncated sfnt directory')
  const flavor = u32(buffer, faceOffset)
  if (!isSfntFlavor(flavor)) throw new TypeError('unsupported sfnt flavor')
  const count = u16(buffer, faceOffset + 4)
  if (count === 0 || count > MAX_TABLES || !contains(buffer, faceOffset + 12, count * 16)) {
    throw new RangeError('invalid sfnt table count')
  }

  const records: TableRecord[] = []
  const seen = new Set<string>()
  for (let i = 0; i < count; i++) {
    const cursor = faceOffset + 12 + i * 16
    const tag = tagAt(buffer, cursor)
    const record = {
      tag,
      checksum: u32(buffer, cursor + 4),
      offset: u32(buffer, cursor + 8),
      length: u32(buffer, cursor + 12),
    }
    if (seen.has(tag) || !contains(buffer, record.offset, record.length)) throw new RangeError('invalid sfnt table record')
    seen.add(tag)
    records.push(record)
  }
  return { flavor, records }
}

function decodeUtf16Be(bytes: Buffer): string {
  if ((bytes.length & 1) !== 0) return ''
  const codeUnits = new Uint16Array(bytes.length / 2)
  for (let i = 0; i < codeUnits.length; i++) codeUnits[i] = bytes.readUInt16BE(i * 2)
  let result = ''
  for (let i = 0; i < codeUnits.length; i += 2048) {
    result += String.fromCharCode(...codeUnits.subarray(i, i + 2048))
  }
  return result
}

function cleanName(value: string): string {
  return value.replace(/\0/g, '').replace(/\s+/g, ' ').trim()
}

interface DecodedName {
  id: number
  value: string
  preference: number
  order: number
}

function decodeNames(buffer: Buffer, record: TableRecord): FaceNames {
  if (record.length < 6) throw new RangeError('truncated name table')
  const base = record.offset
  const format = u16(buffer, base)
  const count = u16(buffer, base + 2)
  const storageOffset = u16(buffer, base + 4)
  if ((format !== 0 && format !== 1) || count > 0xffff || 6 + count * 12 > record.length || storageOffset > record.length) {
    throw new RangeError('invalid name table header')
  }

  const decoded: DecodedName[] = []
  for (let i = 0; i < count; i++) {
    const cursor = base + 6 + i * 12
    const platformId = u16(buffer, cursor)
    const encodingId = u16(buffer, cursor + 2)
    const languageId = u16(buffer, cursor + 4)
    const nameId = u16(buffer, cursor + 6)
    const length = u16(buffer, cursor + 8)
    const relativeOffset = u16(buffer, cursor + 10)
    if (![1, 2, 6, 16, 17].includes(nameId)) continue
    const start = base + storageOffset + relativeOffset
    if (relativeOffset > record.length - storageOffset || length > record.length - storageOffset - relativeOffset || !contains(buffer, start, length)) continue

    const bytes = buffer.subarray(start, start + length)
    let value = ''
    let preference = 0
    if (platformId === 0 || platformId === 3) {
      value = decodeUtf16Be(bytes)
      preference = platformId === 3 && languageId === 0x0409 ? 40 : platformId === 0 ? 30 : 20
      if (platformId === 3 && ![0, 1, 10].includes(encodingId)) preference -= 10
    } else if (platformId === 1) {
      value = bytes.toString('latin1')
      preference = languageId === 0 ? 10 : 5
    }
    value = cleanName(value)
    if (value) decoded.push({ id: nameId, value, preference, order: i })
  }

  const values = (...ids: number[]): string[] => {
    const candidates = decoded
      .filter((item) => ids.includes(item.id))
      .sort((left, right) => ids.indexOf(left.id) - ids.indexOf(right.id) || right.preference - left.preference || left.order - right.order)
    const output: string[] = []
    const keys = new Set<string>()
    for (const candidate of candidates) {
      const key = norm(candidate.value)
      if (!key || keys.has(key)) continue
      keys.add(key)
      output.push(candidate.value)
    }
    return output
  }

  return {
    ps: values(6),
    families: values(16, 1),
    subfamilies: values(17, 2),
  }
}

/** Normalize a font name for deterministic identity lookup. */
export const norm = (s: string): string => s
  .normalize('NFKD')
  .toLocaleLowerCase('en-US')
  .replace(/\p{Mark}+/gu, '')
  .replace(/[^a-z0-9]+/g, '')

const TOKEN_ALIASES: ReadonlyArray<readonly [string, string]> = [
  ['ultracondensed', 'ultracondensed'],
  ['extracondensed', 'extracondensed'],
  ['semicondensed', 'semicondensed'],
  ['condensed', 'condensed'],
  ['narrow', 'condensed'],
  ['semiexpanded', 'semiexpanded'],
  ['extraexpanded', 'extraexpanded'],
  ['ultraexpanded', 'ultraexpanded'],
  ['expanded', 'expanded'],
  ['extended', 'expanded'],
  ['extralight', 'extralight'],
  ['ultralight', 'extralight'],
  ['semilight', 'light'],
  ['demilight', 'light'],
  ['light', 'light'],
  ['medium', 'medium'],
  ['semibold', 'semibold'],
  ['demibold', 'semibold'],
  ['extrabold', 'extrabold'],
  ['ultrabold', 'extrabold'],
  ['bold', 'bold'],
  ['black', 'black'],
  ['heavy', 'black'],
  ['thin', 'thin'],
  ['hairline', 'thin'],
  ['italic', 'italic'],
  ['oblique', 'italic'],
  ['regular', 'regular'],
  ['normal', 'regular'],
  ['roman', 'regular'],
  ['book', 'regular'],
]

/** Extract canonical weight, width, and slant words from a style or PostScript name. */
export function styleTokens(ps: string): string[] {
  const remaining = ps
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .toLocaleLowerCase('en-US')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
  const compact = remaining.replace(/\s+/g, '')
  const found: string[] = []

  // Longest spellings run first so “extra bold” is not reduced to “bold”.
  let searchable = compact
  for (const [spelling, canonical] of TOKEN_ALIASES) {
    if (!searchable.includes(spelling)) continue
    if (!found.includes(canonical)) found.push(canonical)
    searchable = searchable.replace(spelling, '')
  }
  return found
}

function comparableStyle(tokens: readonly string[]): Set<string> {
  const canonical = new Set(tokens.flatMap(styleTokens))
  if (canonical.size === 0) canonical.add('regular')
  return canonical
}

/** Higher scores are closer; extra or missing style coordinates reduce a match. */
export function styleScore(face: FaceRef, want: string[]): number {
  const actual = comparableStyle([face.style])
  const desired = comparableStyle(want)
  let score = 0
  for (const token of desired) score += actual.has(token) ? 1000 : -1000
  for (const token of actual) if (!desired.has(token)) score -= 1
  return score
}

/** @internal Parse independently supplied sfnt bytes for inventory construction. */
export function fontFacesFromBuffer(buffer: Buffer, path: string): ParsedFace[] {
  if (buffer.length > MAX_FONT_BYTES) return []
  const faces: ParsedFace[] = []
  try {
    for (const offset of faceOffsets(buffer)) {
      const directory = tableDirectory(buffer, offset)
      const name = directory.records.find((item) => item.tag === 'name')
      if (!name) continue
      const names = decodeNames(buffer, name)
      if (names.ps.length === 0 && names.families.length === 0) continue
      faces.push({
        ref: { path, offset, style: names.subfamilies[0] ?? 'Regular' },
        names,
      })
    }
  } catch {
    return []
  }
  return faces
}

function searchRoots(): string[] {
  const home = homedir()
  if (platform() === 'darwin') return [join(home, 'Library/Fonts'), '/Library/Fonts', '/System/Library/Fonts', '/Network/Library/Fonts']
  if (platform() === 'win32') {
    const windows = process.env.WINDIR ?? 'C:\\Windows'
    const local = process.env.LOCALAPPDATA
    const roaming = process.env.APPDATA
    return [join(windows, 'Fonts'), local ? join(local, 'Microsoft/Windows/Fonts') : '', roaming ? join(roaming, 'Microsoft/Windows/Fonts') : ''].filter(Boolean)
  }
  const dataHome = process.env.XDG_DATA_HOME || join(home, '.local/share')
  const dataDirs = (process.env.XDG_DATA_DIRS || '/usr/local/share:/usr/share').split(':')
  return [join(home, '.fonts'), join(dataHome, 'fonts'), ...dataDirs.map((directory) => join(directory, 'fonts'))]
}

function discoverFontFiles(): string[] {
  const found: string[] = []
  const seenDirectories = new Set<string>()
  const visit = (directory: string, depth: number) => {
    if (!directory || depth > 12 || found.length >= 100_000 || seenDirectories.has(directory) || !existsSync(directory)) return
    seenDirectories.add(directory)
    let entries
    try {
      entries = readdirSync(directory, { withFileTypes: true }).sort((a, b) => a.name.localeCompare(b.name, 'en'))
    } catch {
      return
    }
    for (const entry of entries) {
      const path = join(directory, entry.name)
      if (entry.isDirectory()) visit(path, depth + 1)
      else if (entry.isFile() && FONT_EXTENSIONS.has(extname(entry.name).toLocaleLowerCase('en-US'))) found.push(path)
    }
  }
  for (const root of searchRoots()) visit(root, 0)
  return [...new Set(found)].sort((a, b) => a.localeCompare(b, 'en'))
}

export function resetFontIndexForTests(): void {
  cachedIndex = undefined
}

export function getFontIndex(): FontIndex {
  if (cachedIndex) return cachedIndex
  const byPs = new Map<string, FaceRef>()
  const byFamily = new Map<string, FaceRef[]>()

  for (const path of discoverFontFiles()) {
    let buffer: Buffer
    try {
      const size = statSync(path).size
      if (size <= 0 || size > MAX_FONT_BYTES) continue
      buffer = readFileSync(path)
    } catch {
      continue
    }
    for (const face of fontFacesFromBuffer(buffer, path)) {
      for (const ps of face.names.ps) {
        const key = norm(ps)
        if (key && !byPs.has(key)) byPs.set(key, face.ref)
      }
      for (const family of face.names.families) {
        const key = norm(family)
        if (!key) continue
        const refs = byFamily.get(key) ?? []
        if (!refs.some((item) => item.path === face.ref.path && item.offset === face.ref.offset)) refs.push(face.ref)
        byFamily.set(key, refs)
      }
    }
  }

  cachedIndex = { byPs, byFamily }
  return cachedIndex
}

function readAt(fd: number, position: number, length: number): Buffer {
  const output = Buffer.alloc(length)
  let filled = 0
  while (filled < length) {
    const count = readSync(fd, output, filled, length - filled, position + filled)
    if (count === 0) throw new RangeError('truncated font file')
    filled += count
  }
  return output
}

function cmapTableFromFd(fd: number, faceOffset: number): Buffer {
  const header = readAt(fd, faceOffset, 12)
  if (!isSfntFlavor(u32(header, 0))) throw new TypeError('unsupported sfnt flavor')
  const count = u16(header, 4)
  if (count === 0 || count > MAX_TABLES) throw new RangeError('invalid sfnt table count')
  const records = readAt(fd, faceOffset + 12, count * 16)
  for (let i = 0; i < count; i++) {
    const cursor = i * 16
    if (tagAt(records, cursor) !== 'cmap') continue
    const offset = u32(records, cursor + 8)
    const length = u32(records, cursor + 12)
    const size = fstatSync(fd).size
    if (length < 4 || length > 64 * 1024 * 1024 || offset > size - length) throw new RangeError('invalid cmap table')
    return readAt(fd, offset, length)
  }
  throw new RangeError('font has no cmap table')
}

function cmapRank(platformId: number, encodingId: number, format: number): number {
  if (format === 12 || format === 13) {
    if (platformId === 3 && encodingId === 10) return 500
    if (platformId === 0 && (encodingId === 4 || encodingId === 6)) return 490
    if (platformId === 0) return 450
  }
  if (format === 4) {
    if (platformId === 3 && encodingId === 1) return 400
    if (platformId === 0) return 390
  }
  return -1
}

function cmapLookupFormat4(table: Buffer, base: number, cp: number): number {
  if (cp > 0xffff || !contains(table, base, 16)) return 0
  const length = u16(table, base + 2)
  const segCountX2 = u16(table, base + 6)
  if ((segCountX2 & 1) !== 0 || segCountX2 === 0 || !contains(table, base, length)) return 0
  const count = segCountX2 / 2
  const endBase = base + 14
  const startBase = endBase + count * 2 + 2
  const deltaBase = startBase + count * 2
  const rangeBase = deltaBase + count * 2
  if (!contains(table, rangeBase, count * 2)) return 0

  let low = 0
  let high = count - 1
  while (low <= high) {
    const mid = (low + high) >>> 1
    const end = u16(table, endBase + mid * 2)
    if (cp > end) low = mid + 1
    else high = mid - 1
  }
  if (low >= count) return 0
  const start = u16(table, startBase + low * 2)
  if (cp < start) return 0
  const delta = u16(table, deltaBase + low * 2)
  const range = u16(table, rangeBase + low * 2)
  if (range === 0) return (cp + delta) & 0xffff
  const glyphOffset = rangeBase + low * 2 + range + (cp - start) * 2
  if (!contains(table, glyphOffset, 2) || glyphOffset >= base + length) return 0
  const glyph = u16(table, glyphOffset)
  return glyph === 0 ? 0 : (glyph + delta) & 0xffff
}

function cmapLookupGroups(table: Buffer, base: number, cp: number, format: 12 | 13): number {
  if (!contains(table, base, 16)) return 0
  const length = u32(table, base + 4)
  const count = u32(table, base + 12)
  if (count > 0x100000 || length < 16 || !contains(table, base, length) || !contains(table, base + 16, count * 12)) return 0
  let low = 0
  let high = count - 1
  while (low <= high) {
    const mid = (low + high) >>> 1
    const cursor = base + 16 + mid * 12
    const start = u32(table, cursor)
    const end = u32(table, cursor + 4)
    if (cp < start) high = mid - 1
    else if (cp > end) low = mid + 1
    else {
      const glyph = u32(table, cursor + 8)
      return format === 12 ? (glyph + cp - start) >>> 0 : glyph
    }
  }
  return 0
}

function bestUnicodeCmap(table: Buffer): { base: number; format: 4 | 12 | 13 } | undefined {
  if (u16(table, 0) !== 0) return undefined
  const count = u16(table, 2)
  if (!contains(table, 4, count * 8)) return undefined
  let best: { base: number; format: 4 | 12 | 13; rank: number } | undefined
  for (let i = 0; i < count; i++) {
    const cursor = 4 + i * 8
    const platformId = u16(table, cursor)
    const encodingId = u16(table, cursor + 2)
    const base = u32(table, cursor + 4)
    if (!contains(table, base, 2)) continue
    const format = u16(table, base)
    if (format !== 4 && format !== 12 && format !== 13) continue
    const rank = cmapRank(platformId, encodingId, format)
    if (rank >= 0 && (!best || rank > best.rank)) best = { base, format, rank }
  }
  return best
}

export function faceCoversCodepoints(fd: number, faceOffset: number, codepoints: readonly number[]): boolean {
  try {
    const cmap = cmapTableFromFd(fd, faceOffset)
    const subtable = bestUnicodeCmap(cmap)
    if (!subtable) return false
    return codepoints.every((cp) => {
      if (!Number.isInteger(cp) || cp < 0 || cp > 0x10ffff || (cp >= 0xd800 && cp <= 0xdfff)) return false
      const glyph = subtable.format === 4
        ? cmapLookupFormat4(cmap, subtable.base, cp)
        : cmapLookupGroups(cmap, subtable.base, cp, subtable.format)
      return glyph !== 0
    })
  } catch {
    return false
  }
}

function checksum(buffer: Buffer): number {
  let sum = 0
  for (let offset = 0; offset < buffer.length; offset += 4) {
    let word = 0
    for (let byte = 0; byte < 4; byte++) word = (word << 8) | (buffer[offset + byte] ?? 0)
    sum = (sum + (word >>> 0)) >>> 0
  }
  return sum
}

/** @internal Return one collection face as a self-contained sfnt buffer. */
export function readFaceBytes(ref: FaceRef): Buffer | null {
  let source: Buffer
  try {
    source = readFileSync(ref.path)
    if (source.length > MAX_FONT_BYTES) return null
    if (ref.offset === 0 && u32(source, 0) !== TTC_TAG) {
      tableDirectory(source, 0)
      return Buffer.from(source)
    }
    if (!faceOffsets(source).includes(ref.offset)) return null
    const { flavor, records } = tableDirectory(source, ref.offset)
    const directorySize = 12 + records.length * 16
    let outputSize = directorySize
    for (const record of records) outputSize += (record.length + 3) & ~3
    if (outputSize > MAX_FONT_BYTES) return null
    const output = Buffer.alloc(outputSize)
    output.writeUInt32BE(flavor, 0)
    output.writeUInt16BE(records.length, 4)
    const power = 2 ** Math.floor(Math.log2(records.length))
    output.writeUInt16BE(power * 16, 6)
    output.writeUInt16BE(Math.log2(power), 8)
    output.writeUInt16BE(records.length * 16 - power * 16, 10)
    let dataOffset = directorySize
    let headOffset: number | undefined
    records.forEach((record, index) => {
      const cursor = 12 + index * 16
      output.write(record.tag, cursor, 4, 'latin1')
      output.writeUInt32BE(record.checksum, cursor + 4)
      output.writeUInt32BE(dataOffset, cursor + 8)
      output.writeUInt32BE(record.length, cursor + 12)
      source.copy(output, dataOffset, record.offset, record.offset + record.length)
      if (record.tag === 'head' && record.length >= 12) headOffset = dataOffset
      dataOffset += (record.length + 3) & ~3
    })
    if (headOffset !== undefined) {
      output.writeUInt32BE(0, headOffset + 8)
      output.writeUInt32BE((0xb1b0afba - checksum(output)) >>> 0, headOffset + 8)
    }
    return output
  } catch {
    return null
  }
}

export function isTruetype(bytes: Buffer): boolean {
  try {
    if (bytes.length < 4) return false
    const signature = u32(bytes, 0)
    if (signature === SFNT_1 || signature === TRUE_TAG) return true
    if (signature === OTTO_TAG) return false
    const offset = faceOffsets(bytes)[0]!
    const { flavor, records } = tableDirectory(bytes, offset)
    return flavor !== OTTO_TAG && records.some((record) => record.tag === 'glyf') && records.some((record) => record.tag === 'loca')
  } catch {
    return false
  }
}

/** @internal Iterate unique indexed faces in deterministic lookup order. */
export function indexedFaces(index: FontIndex): FaceRef[] {
  const output: FaceRef[] = []
  const seen = new Set<string>()
  for (const refs of index.byFamily.values()) {
    for (const ref of refs) {
      const key = `${ref.path}\0${ref.offset}`
      if (!seen.has(key)) {
        seen.add(key)
        output.push(ref)
      }
    }
  }
  return output
}

/** @internal Check coverage without retaining an open descriptor. */
export function refCoversCodepoints(ref: FaceRef, codepoints: readonly number[]): boolean {
  let fd: number | undefined
  try {
    fd = openSync(ref.path, 'r')
    return faceCoversCodepoints(fd, ref.offset, codepoints)
  } catch {
    return false
  } finally {
    if (fd !== undefined) closeSync(fd)
  }
}
