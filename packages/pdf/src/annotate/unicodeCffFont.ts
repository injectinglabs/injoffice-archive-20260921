import { validateCffPrograms } from './unicodeCffProgram.js'
import { expertCffCharsets } from './unicodeCffCharsets.js'

export interface UnicodeCffFont { bytes: Uint8Array; glyphCids: readonly number[]; name: string; registry: string; ordering: string; supplement: number }
interface Index { values: Uint8Array[]; start: number; end: number }
interface Operand { values: number[]; bytes: Uint8Array }
type Dict = Map<number, Operand>
const maxBytes = 16 * 1024 * 1024
const encode = new TextEncoder()
const join = (...parts: Uint8Array[]) => {
  const length = parts.reduce((sum, bytes) => sum + bytes.length, 0)
  if (length > maxBytes) throw new Error('CFF embedding byte limit exceeded')
  const output = new Uint8Array(length)
  let offset = 0
  for (const bytes of parts) { output.set(bytes, offset); offset += bytes.length }
  return output
}
const number = (n: number) => new Uint8Array([29, (n >>> 24) & 255, (n >>> 16) & 255, (n >>> 8) & 255, n & 255])
const entry = (operator: number, ...values: number[]) => join(...values.map(number), operator >= 1200 ? new Uint8Array([12, operator - 1200]) : new Uint8Array([operator]))
function index(values: readonly Uint8Array[]): Uint8Array {
  if (values.length === 0) return new Uint8Array(2)
  if (values.length > 65535) throw new Error('CFF INDEX count exceeded')
  const header = new Uint8Array(3 + (values.length + 1) * 4), view = new DataView(header.buffer)
  view.setUint16(0, values.length); header[2] = 4
  let offset = 1
  values.forEach((value, i) => { view.setUint32(3 + i * 4, offset); offset += value.length })
  view.setUint32(3 + values.length * 4, offset)
  return join(header, ...values)
}

/** Validates CFF1 data and creates a CID-keyed program without rewriting glyphs,
 * hint data or subroutines. Source CIDs survive; name-keyed faces get GID CIDs.
 * Adobe CFF 5176 §§4–6, 13–19; the Type 2 program preflight is kept separate. */
export function prepareUnicodeCffFont(source: Uint8Array, glyphCount: number, unitsPerEm: number): UnicodeCffFont {
  if (source.length < 4 || source.length > maxBytes || !Number.isInteger(glyphCount) || glyphCount < 1 || glyphCount > 65535) throw new Error('invalid CFF font size or glyph count')
  const view = new DataView(source.buffer, source.byteOffset, source.byteLength)
  const range = (offset: number, length: number) => { if (!Number.isSafeInteger(offset) || !Number.isSafeInteger(length) || offset < 0 || length < 0 || offset + length > source.length) throw new Error('CFF range exceeds font bytes') }
  const byte = (offset: number) => { range(offset, 1); return source[offset]! }
  const u16 = (offset: number) => { range(offset, 2); return view.getUint16(offset) }
  const indexes = new Map<number, Index>()
  let indexEntries = 0
  const readIndex = (start: number): Index => {
    const cached = indexes.get(start)
    if (cached) return cached
    const count = u16(start)
    indexEntries += count
    if (indexEntries > 262144) throw new Error('CFF INDEX entry limit exceeded')
    if (!count) { const result = { start, end: start + 2, values: [] }; indexes.set(start, result); return result }
    const size = byte(start + 2)
    if (size < 1 || size > 4) throw new Error('invalid CFF INDEX offSize')
    const data = start + 3 + (count + 1) * size
    range(start, data - start)
    const offsets: number[] = []
    for (let i = 0; i <= count; i++) {
      let value = 0
      for (let j = 0; j < size; j++) value = value * 256 + byte(start + 3 + i * size + j)
      if (value < 1 || (i === 0 && value !== 1) || i > 0 && value < offsets[i - 1]!) throw new Error('invalid CFF INDEX offsets')
      offsets.push(value)
    }
    range(data, offsets[count]! - 1)
    const result = { start, end: data + offsets[count]! - 1, values: offsets.slice(0, count).map((offset, i) => source.subarray(data + offset - 1, data + offsets[i + 1]! - 1)) }
    indexes.set(start, result)
    return result
  }
  const dict = (bytes: Uint8Array): Dict => {
    const values: number[] = [], result: Dict = new Map()
    const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
    let position = 0, start = 0
    const next = () => { if (position >= bytes.length) throw new Error('truncated CFF DICT'); return bytes[position++]! }
    while (position < bytes.length) {
      const b = next()
      if (b <= 21) {
        const op = b === 12 ? 1200 + next() : b
        if (result.has(op) || values.length === 0) throw new Error('duplicate or empty CFF DICT operator')
        result.set(op, { values: values.splice(0), bytes: bytes.slice(start, position) }); start = position
      } else {
        let n: number
        if (b === 28) { const a = next(), c = next(); n = (a * 256 + c) << 16 >> 16 }
        else if (b === 29) { if (position + 4 > bytes.length) throw new Error('truncated CFF DICT integer'); n = v.getInt32(position); position += 4 }
        else if (b === 30) {
          let text = '', done = false
          while (!done) {
            const pair = next()
            for (const nibble of [pair >>> 4, pair & 15]) {
              if (nibble === 15) { done = true; break }
              if (nibble === 13 || text.length > 64) throw new Error('invalid CFF real operand')
              text += nibble <= 9 ? String(nibble) : nibble === 10 ? '.' : nibble === 11 ? 'E' : nibble === 12 ? 'E-' : '-'
            }
          }
          if (!/^-?(?:\d+(?:\.\d*)?|\.\d+)(?:E-?\d+)?$/.test(text)) throw new Error('invalid CFF real operand')
          n = Number(text)
        } else if (b >= 32 && b <= 246) n = b - 139
        else if (b >= 247 && b <= 250) n = (b - 247) * 256 + next() + 108
        else if (b >= 251 && b <= 254) n = -(b - 251) * 256 - next() - 108
        else throw new Error('reserved CFF DICT operand')
        if (!Number.isFinite(n) || values.length >= 48) throw new Error('invalid CFF DICT operand')
        values.push(n)
      }
    }
    if (values.length) throw new Error('trailing CFF DICT operands')
    return result
  }
  const operands = (d: Dict, op: number, count: number, fallback?: number[]): number[] => {
    const values = d.get(op)?.values ?? fallback
    if (!values || values.length !== count) throw new Error('invalid or missing CFF DICT operands')
    return values
  }
  const integer = (n: number) => { if (!Number.isSafeInteger(n) || n < 0 || n > 0x7fffffff) throw new Error('invalid CFF integer'); return n }
  const offset = (d: Dict, op: number) => integer(operands(d, op, 1)[0]!)
  if (byte(0) !== 1 || byte(1) !== 0 || byte(2) < 4 || byte(3) < 1 || byte(3) > 4) throw new Error('unsupported CFF header')
  range(0, byte(2))
  const names = readIndex(byte(2)), tops = readIndex(names.end), strings = readIndex(tops.end), globals = readIndex(strings.end)
  if (names.values.length !== 1 || tops.values.length !== 1 || strings.values.length > 64607) throw new Error('CFF must contain one bounded font')
  const name = names.values[0]!
  if (name.length < 1 || name.length > 127 || [...name].some(b => b < 33 || b > 126 || '[](){}<>/%'.includes(String.fromCharCode(b)))) throw new Error('invalid CFF font name')
  const top = dict(tops.values[0]!)
  const isCid = top.has(1230)
  const validSid = (sid: number) => Number.isInteger(sid) && sid >= 0 && sid < 391 + strings.values.length
  const checkDict = (d: Dict, kind: 'top' | 'fd' | 'private') => {
    const sizes: Record<number, number | undefined> = kind === 'top'
      ? { 0: 1, 1: 1, 2: 1, 3: 1, 4: 1, 5: 4, 13: 1, 14: -1, 15: 1, 16: 1, 17: 1, 18: 2, 1200: 1, 1201: 1, 1202: 1, 1203: 1, 1204: 1, 1205: 1, 1206: 1, 1207: 6, 1208: 1, 1222: 1, ...(isCid ? { 1230: 3, 1231: 1, 1232: 1, 1233: 1, 1234: 1, 1235: 1, 1236: 1, 1237: 1 } : {}) }
      : kind === 'fd' ? { 18: 2, 1207: 6, 1238: 1 }
        : { 6: -1, 7: -1, 8: -1, 9: -1, 10: 1, 11: 1, 19: 1, 20: 1, 21: 1, 1209: 1, 1210: 1, 1211: 1, 1212: -1, 1213: -1, 1214: 1, 1217: 1, 1218: 1, 1219: 1 }
    for (const [op, { values }] of d) {
      if (sizes[op] === undefined || sizes[op] !== -1 && sizes[op] !== values.length) throw new Error('unsupported or invalid CFF DICT operator')
      const sidOperator = kind === 'top' && [0, 1, 2, 3, 4, 1200, 1222].includes(op) || kind === 'fd' && op === 1238
      if (sidOperator && !validSid(values[0]!)) throw new Error('CFF DICT references missing SID')
      if (kind === 'private' && [6, 7, 8, 9].includes(op) && (values.length % 2 || values.length > ([6, 8].includes(op) ? 14 : 10))) throw new Error('invalid CFF blue zone array')
      if (kind === 'private' && [1212, 1213].includes(op) && values.length > 12) throw new Error('invalid CFF stem snap array')
      if ((kind === 'private' && op === 1214 || kind === 'top' && op === 1201) && ![0, 1].includes(values[0]!)) throw new Error('invalid CFF boolean')
    }
  }
  checkDict(top, 'top')
  if (isCid && top.keys().next().value !== 1230) throw new Error('CFF ROS must be first')
  if ([1220, 1221, 1223].some(op => top.has(op)) || operands(top, 1205, 1, [0])[0] !== 0 || operands(top, 1206, 1, [2])[0] !== 2 || operands(top, 1233, 1, [0])[0] !== 0) throw new Error('unsupported synthetic, stroked or non-Type2 CFF font')
  const matrix = operands(top, 1207, 6, [0.001, 0, 0, 0.001, 0, 0])
  const expected = [1 / unitsPerEm, 0, 0, 1 / unitsPerEm, 0, 0]
  if (!Number.isFinite(unitsPerEm) || unitsPerEm <= 0 || matrix.some((n, i) => Math.abs(n - expected[i]!) > 1e-12)) throw new Error('CFF FontMatrix must match horizontal OpenType units')
  const charstringsOffset = offset(top, 17), glyphs = readIndex(charstringsOffset)
  if (glyphs.values.length !== glyphCount || glyphs.values.some(bytes => bytes.length === 0)) throw new Error('CFF and OpenType glyph counts disagree')
  const occupied: { start: number; end: number; kind: string }[] = [{ start: 0, end: globals.end, kind: 'header' }, { start: glyphs.start, end: glyphs.end, kind: 'glyphs' }]
  const addRange = (start: number, length: number, kind: string) => { range(start, length); if (length) occupied.push({ start, end: start + length, kind }) }
  const charsetOffset = integer(operands(top, 15, 1, [0])[0]!), glyphCids = [0]
  if (charsetOffset <= 2) {
    if (isCid || glyphCount > [229, 166, 87][charsetOffset]!) throw new Error('invalid predefined CFF charset')
    for (let gid = 1; gid < glyphCount; gid++) glyphCids.push(charsetOffset === 0 ? gid : expertCffCharsets[charsetOffset - 1]![gid]!)
  } else {
    let position = charsetOffset
    const format = byte(position++)
    if (format > 2) throw new Error('invalid CFF charset format')
    while (glyphCids.length < glyphCount) {
      const first = u16(position); position += 2
      const count = format === 0 ? 1 : (format === 1 ? byte(position++) : u16(position)) + 1
      if (format === 2) position += 2
      if (first < 1 || first + count > (isCid ? 65536 : 391 + strings.values.length) || glyphCids.length + count > glyphCount) throw new Error('invalid CFF charset range or missing SID')
      for (let i = 0; i < count; i++) glyphCids.push(first + i)
    }
    if (new Set(glyphCids).size !== glyphCount) throw new Error('duplicate CFF charset entries')
    addRange(charsetOffset, position - charsetOffset, 'charset')
  }
  const privateData = (d: Dict): { size: number; start: number; locals: Uint8Array[] } => {
    const [size, start] = operands(d, 18, 2, [0, 0]).map(integer) as [number, number]
    addRange(start, size, 'private')
    const parsed = dict(source.subarray(start, start + size))
    checkDict(parsed, 'private')
    let localValues: Uint8Array[] = []
    if (parsed.has(19)) {
      const relative = offset(parsed, 19)
      if (relative < size) throw new Error('CFF local subroutines overlap Private DICT')
      const local = readIndex(start + relative); addRange(local.start, local.end - local.start, 'subrs'); localValues = local.values
    }
    return { size, start, locals: localValues }
  }
  let registry = 'Adobe', ordering = 'Identity', supplement = 0
  const selected = new Array<number>(glyphCount).fill(0), locals: Uint8Array[][] = []
  let originalPrivate: ReturnType<typeof privateData> | undefined
  if (isCid) {
    if (top.has(16) || top.has(18)) throw new Error('CID CFF must use FDArray private dictionaries without Encoding')
    const ros = operands(top, 1230, 3)
    const string = (sid: number) => {
      integer(sid)
      // ROS names in practical CID fonts are custom String INDEX entries.
      const bytes = strings.values[sid - 391]
      if (!bytes || !bytes.length || bytes.length > 127 || [...bytes].some(b => b < 32 || b > 126)) throw new Error('unsupported or invalid CFF ROS string')
      return String.fromCharCode(...bytes)
    }
    registry = string(ros[0]!); ordering = string(ros[1]!); supplement = integer(ros[2]!)
    const count = integer(operands(top, 1234, 1, [8720])[0]!)
    if (count > 65536 || Math.max(...glyphCids) >= count) throw new Error('CFF CIDCount does not cover charset')
    const fds = readIndex(offset(top, 1236))
    if (!fds.values.length || fds.values.length > 256) throw new Error('invalid CFF FDArray count')
    addRange(fds.start, fds.end - fds.start, 'fdarray')
    for (const fd of fds.values) {
      const d = dict(fd), fdMatrix = operands(d, 1207, 6, [1, 0, 0, 1, 0, 0])
      checkDict(d, 'fd')
      if (fdMatrix.some((n, i) => n !== [1, 0, 0, 1, 0, 0][i])) throw new Error('unsupported CFF FD FontMatrix composition')
      locals.push(privateData(d).locals)
    }
    const start = offset(top, 1237)
    let position = start
    const format = byte(position++)
    if (format === 0) for (let gid = 0; gid < glyphCount; gid++) selected[gid] = byte(position++)
    else if (format === 3) {
      const count = u16(position); position += 2
      if (!count || count > glyphCount) throw new Error('invalid CFF FDSelect ranges')
      let previous = 0, fd = 0
      for (let i = 0; i < count; i++) {
        const first = u16(position), nextFd = byte(position + 2); position += 3
        if (i === 0 ? first !== 0 : first <= previous || first >= glyphCount) throw new Error('invalid CFF FDSelect ordering')
        if (i) selected.fill(fd, previous, first)
        previous = first; fd = nextFd
      }
      if (u16(position) !== glyphCount) throw new Error('invalid CFF FDSelect sentinel')
      position += 2; selected.fill(fd, previous)
    } else throw new Error('unsupported CFF FDSelect format')
    if (selected.some(fd => fd >= fds.values.length)) throw new Error('invalid CFF FDSelect index')
    addRange(start, position - start, 'fdselect')
  } else {
    if (top.has(1236) || top.has(1237)) throw new Error('name-keyed CFF has CID-only dictionaries')
    originalPrivate = privateData(top); locals.push(originalPrivate.locals)
    // OpenType uses cmap rather than CFF Encoding; still validate its ranges.
    const encoding = integer(operands(top, 16, 1, [0])[0]!)
    if (encoding > 1) {
      let position = encoding
      const format = byte(position++), count = byte(position++), codes = new Set<number>()
      let assigned = 0
      if ((format & 127) > 1) throw new Error('invalid CFF Encoding format')
      for (let i = 0; i < count; i++) {
        const first = byte(position++), amount = (format & 127) === 0 ? 1 : byte(position++) + 1
        if (first + amount > 256) throw new Error('invalid CFF Encoding range')
        for (let j = 0; j < amount; j++) { if (codes.has(first + j)) throw new Error('duplicate CFF Encoding code'); codes.add(first + j) }
        assigned += amount
      }
      if (assigned >= glyphCount) throw new Error('invalid CFF Encoding glyph count')
      if (format & 128) { const supplements = byte(position++); for (let i = 0; i < supplements; i++) { const code = byte(position++), sid = u16(position); position += 2; if (codes.has(code) || sid < 1 || sid >= 391 + strings.values.length || !glyphCids.slice(1, assigned + 1).includes(sid)) throw new Error('invalid CFF Encoding supplement'); codes.add(code) } }
      addRange(encoding, position - encoding, 'encoding')
    }
  }
  occupied.sort((a, b) => a.start - b.start)
  let end = 0, previous: typeof occupied[number] | undefined
  for (const current of occupied) {
    // Shared private/subroutine objects across FD entries are valid; partial
    // overlaps and structures that alias unrelated metadata are not.
    if (current.start < end && !(previous?.start === current.start && previous.end === current.end && previous.kind === current.kind && ['private', 'subrs'].includes(current.kind))) throw new Error('overlapping CFF structures')
    end = Math.max(end, current.end); previous = current
  }
  validateCffPrograms(glyphs.values, globals.values, locals, selected)
  if (isCid) return { bytes: source.slice(), glyphCids, name: String.fromCharCode(...name), registry, ordering, supplement }

  // Preserve original Type 2 bytes and Private-relative local Subrs offsets in
  // an opaque relocated payload. Only new absolute Top/FD pointers change.
  const charset = new Uint8Array(glyphCount > 1 ? 5 : 1)
  charset[0] = 2
  if (glyphCount > 1) { new DataView(charset.buffer).setUint16(1, 1); new DataView(charset.buffer).setUint16(3, glyphCount - 2) }
  const fdSelect = new Uint8Array([3, 0, 1, 0, 0, 0, glyphCount >>> 8, glyphCount & 255])
  const stringsOut = index([...strings.values, encode.encode('Adobe'), encode.encode('Identity')])
  const namesOut = index(names.values), globalOut = source.slice(globals.start, globals.end)
  const retained = [...top].filter(([op]) => ![15, 16, 17, 18].includes(op)).map(([, operand]) => operand.bytes)
  const makeTop = (charsetOffset: number, fdSelectOffset: number, fdArrayOffset: number, payload: number) => index([join(entry(1230, 391 + strings.values.length, 392 + strings.values.length, 0), ...retained, entry(1234, glyphCount), entry(15, charsetOffset), entry(17, payload + charstringsOffset), entry(1236, fdArrayOffset), entry(1237, fdSelectOffset))])
  const makeFd = (payload: number) => index([entry(18, originalPrivate!.size, payload + originalPrivate!.start)])
  const topLength = makeTop(0, 0, 0, 0).length, fdLength = makeFd(0).length
  const charsetStart = 4 + namesOut.length + topLength + stringsOut.length + globalOut.length
  const fdSelectStart = charsetStart + charset.length, fdArrayStart = fdSelectStart + fdSelect.length, payload = fdArrayStart + fdLength
  if (payload + source.length > maxBytes) throw new Error('converted CFF exceeds embedded byte limit')
  const bytes = join(new Uint8Array([1, 0, 4, 4]), namesOut, makeTop(charsetStart, fdSelectStart, fdArrayStart, payload), stringsOut, globalOut, charset, fdSelect, makeFd(payload), source)
  return { bytes, glyphCids: Array.from({ length: glyphCount }, (_, gid) => gid), name: String.fromCharCode(...name), registry, ordering, supplement }
}
