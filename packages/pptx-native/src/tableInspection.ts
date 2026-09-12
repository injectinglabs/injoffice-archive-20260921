import type { NativePptxDeck } from './types'
import { assertNativePptx } from './validate'

export interface NativePptxInspectionRect { x: number; y: number; width: number; height: number }
export interface NativePptxInspectedCell { row: number; column: number; rect: NativePptxInspectionRect; paragraphs: string[] }
export interface NativePptxInspectedTable {
  slide_id: string; slide_index: number; part_name: string; part_sha256: string; slide_source_sha256: string
  object_id: string; source_sha256: string; rect: NativePptxInspectionRect; cells: NativePptxInspectedCell[]; warnings: string[]
}
export interface NativePptxTableOmission { slide_id: string; object_id: string; reason: string }
/** Read-only source text and stored geometry. Contains no native mutation or preservation capabilities. */
export interface NativePptxTableInspection {
  protocol: 'pptx-table-content-inspection-v1'; package_sha256: string; source_revision: string
  tables: NativePptxInspectedTable[]; omissions: NativePptxTableOmission[]
}

const warnings = [
  'Plain source text and stored cell geometry only; this is not a rendered Office table.',
  'Table styles, borders and fills are omitted, including any authored dash patterns.',
  'Source text styling, margins, alignment, spacing and inherited font metrics are not applied.',
]
const fail = (): never => { throw new TypeError('Invalid or unjoined PPTX table content inspection.') }
function object(value: unknown, keys: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Object.getPrototypeOf(value) !== Object.prototype) return fail()
  const actual = Reflect.ownKeys(value)
  const expected = keys.split(' ')
  if (actual.length !== expected.length || actual.some(k => typeof k !== 'string' || !expected.includes(k))) return fail()
  for (const k of expected) if (!('value' in Object.getOwnPropertyDescriptor(value, k)!)) return fail()
  return value as Record<string, unknown>
}
function array(value: unknown, max: number): unknown[] {
  if (!Array.isArray(value) || Object.getPrototypeOf(value) !== Array.prototype || value.length > max) return fail()
  const keys = Reflect.ownKeys(value)
  if (keys.length !== value.length + 1) return fail()
  for (let i = 0; i < value.length; i++) {
    const d = Object.getOwnPropertyDescriptor(value, String(i))
    if (!d || !('value' in d)) return fail()
  }
  return value
}
function integer(value: unknown, min: number, max: number): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || Object.is(value, -0) || value < min || value > max) return fail()
  return value
}
function text(value: unknown, max = 4096): string {
  if (typeof value !== 'string' || value.length > max || /[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/u.test(value) || /[\uD800-\uDFFF]/u.test(value)) return fail()
  return value
}
function hash(value: unknown): string {
  if (typeof value !== 'string' || !/^[0-9a-f]{64}$/.test(value)) return fail()
  return value
}
function rect(value: unknown, local: boolean): NativePptxInspectionRect {
  const r = object(value, 'x y width height')
  return { x: integer(r.x, local ? 0 : -1e9, 1e9), y: integer(r.y, local ? 0 : -1e9, 1e9), width: integer(r.width, 1, 1e9), height: integer(r.height, 1, 1e9) }
}

/** Decode against a separately validated strict extraction and an independently computed source-byte SHA-256. */
export function decodeNativePptxTableInspection(input: unknown, deck: NativePptxDeck, packageSHA256: string): NativePptxTableInspection {
  assertNativePptx(deck)
  hash(packageSHA256)
  const v = object(input, 'protocol package_sha256 source_revision tables omissions')
  if (v.protocol !== 'pptx-table-content-inspection-v1' || v.package_sha256 !== packageSHA256 || v.source_revision !== `rev-${packageSHA256}` || deck.origin !== 'parsed' || deck.sourceRevision !== v.source_revision) return fail()
  const tables = array(v.tables, 256), omissions = array(v.omissions, 256)
  if (tables.length + omissions.length > 256) return fail()
  let cells = 0, paragraphs = 0, units = 0
  const ids = new Set<string>(), fingerprints = new Set<string>(), partHashes = new Map<string, string>()
  const decoded = tables.map(value => {
    const t = object(value, 'slide_id slide_index part_name part_sha256 slide_source_sha256 object_id source_sha256 rect cells warnings')
    const slide_index = integer(t.slide_index, 0, deck.slides.length - 1), slide = deck.slides[slide_index]!
    const slide_id = text(t.slide_id, 256), part_name = text(t.part_name, 1024), object_id = text(t.object_id, 32)
    const part_sha256 = hash(t.part_sha256), source_sha256 = hash(t.source_sha256), slide_source_sha256 = hash(t.slide_source_sha256)
    if (slide.id !== slide_id || slide.source?.partName !== part_name || slide.source.fingerprintSha256 !== slide_source_sha256 || !/^cNvPr-[1-9][0-9]{0,9}$/.test(object_id) || Number(object_id.slice(6)) > 4294967295) return fail()
    const elements = slide.elements.filter(e => e.source?.objectId === object_id && e.source.partName === part_name)
    if (elements.length > 1) return fail()
    const element = elements[0]
    if (element ? element.source!.fingerprintSha256 !== source_sha256 : slide.passthrough.filter(p => p.ownerPart === part_name && p.fingerprintSha256 === source_sha256).length !== 1) return fail()
    const key = `${slide_id}\0${object_id}`, fingerprintKey = `${part_name}\0${source_sha256}`
    if (ids.has(key) || fingerprints.has(fingerprintKey) || (partHashes.has(part_name) && partHashes.get(part_name) !== part_sha256)) return fail()
    ids.add(key); fingerprints.add(fingerprintKey); partHashes.set(part_name, part_sha256)
    const frame = rect(t.rect, false)
    const ws = array(t.warnings, 3)
    if (ws.length !== warnings.length || ws.some((w, i) => w !== warnings[i])) return fail()
    const rawCells = array(t.cells, 1024)
    if (!rawCells.length || (cells += rawCells.length) > 4096) return fail()
    const decodedCells = rawCells.map(c => {
      const cell = object(c, 'row column rect paragraphs')
      const ps = array(cell.paragraphs, 256)
      if (!ps.length || (paragraphs += ps.length) > 4096) return fail()
      return { row: integer(cell.row, 0, 31), column: integer(cell.column, 0, 31), rect: rect(cell.rect, true), paragraphs: ps.map(p => {
        const s = text(p, 65536)
        if ((units += s.length) > 65536) return fail()
        return s
      }) }
    })
    const columns = decodedCells.findIndex(c => c.row === 1)
    const width = columns === -1 ? decodedCells.length : columns
    if (width > 32 || decodedCells.length % width !== 0 || decodedCells.length / width > 32) return fail()
    let y = 0
    for (let start = 0; start < decodedCells.length; start += width) {
      let x = 0
      const height = decodedCells[start]!.rect.height
      for (let ci = 0; ci < width; ci++) {
        const c = decodedCells[start + ci]!
        if (c.row !== start / width || c.column !== ci || c.rect.x !== x || c.rect.y !== y || c.rect.height !== height || c.rect.width !== decodedCells[ci]!.rect.width) return fail()
        x += c.rect.width
      }
      if (x !== frame.width) return fail()
      y += height
    }
    if (y !== frame.height) return fail()
    return { slide_id, slide_index, part_name, part_sha256, slide_source_sha256, object_id, source_sha256, rect: frame, cells: decodedCells, warnings: [...warnings] }
  })
  const decodedOmissions = omissions.map(value => {
    const o = object(value, 'slide_id object_id reason')
    const slide_id = text(o.slide_id, 256), object_id = text(o.object_id, 32), reason = text(o.reason)
    if (!deck.slides.some(s => s.id === slide_id) || !reason || !(object_id === 'slide' || object_id === 'unavailable' || /^cNvPr-[1-9][0-9]{0,9}$/.test(object_id))) return fail()
    const key = `${slide_id}\0${object_id}`
    if (object_id !== 'unavailable' && ids.has(key)) return fail()
    if (object_id !== 'unavailable') ids.add(key)
    return { slide_id, object_id, reason }
  })
  return { protocol: 'pptx-table-content-inspection-v1', package_sha256: packageSHA256, source_revision: `rev-${packageSHA256}`, tables: decoded, omissions: decodedOmissions }
}
