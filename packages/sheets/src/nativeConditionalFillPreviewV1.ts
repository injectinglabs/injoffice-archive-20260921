import { snapshotNativePlainData } from './nativePlainData.js'
import type { NativeWorkbookV2 } from './nativeContractV2.generated.js'
import type { NativeWorkbookObjectsV1 } from './nativeObjectsPreviewV1.js'

export type NativeConditionalOperatorV1 = 'equal' | 'notEqual' | 'lessThan' | 'lessThanOrEqual' | 'greaterThan' | 'greaterThanOrEqual'
export interface NativeConditionalFillRuleV1 {
  ref: string
  operator: NativeConditionalOperatorV1
  operand: string
  priority: number
  stop_if_true: boolean
  dxf_id: number
  fill: string
}
export interface NativeConditionalFillCellV1 {
  row: number
  column: number
  lexical: string
  cached: boolean
  matches: boolean
}
export type NativeConditionalFillPreviewV1 = {
  sheet_id: string
  sheet_part: string
  warnings: string[]
} & ({ status: 'available'; rule: NativeConditionalFillRuleV1; cells: NativeConditionalFillCellV1[] } |
  { status: 'unavailable'; rule?: never; cells?: never })

const INTEGER = /^-?(0|[1-9][0-9]{0,14})$/
const OPERATORS: readonly string[] = ['equal', 'notEqual', 'lessThan', 'lessThanOrEqual', 'greaterThan', 'greaterThanOrEqual']
function compare(value: number, operand: number, operator: NativeConditionalOperatorV1): boolean {
  switch (operator) {
    case 'equal': return value === operand
    case 'notEqual': return value !== operand
    case 'lessThan': return value < operand
    case 'lessThanOrEqual': return value <= operand
    case 'greaterThan': return value > operand
    case 'greaterThanOrEqual': return value >= operand
  }
}
function range(ref: string): { row: number; column: number; end_row: number; end_column: number } | undefined {
  const match = /^([A-Z]{1,3})([1-9][0-9]{0,6})(?::([A-Z]{1,3})([1-9][0-9]{0,6}))?$/.exec(ref)
  if (!match) return undefined
  const column = (text: string) => [...text].reduce((value, c) => value * 26 + c.charCodeAt(0) - 64, 0) - 1
  const r = { row: Number(match[2]) - 1, column: column(match[1]!), end_row: Number(match[4] ?? match[2]) - 1, end_column: column(match[3] ?? match[1]!) }
  if (r.row > r.end_row || r.column > r.end_column || r.end_row >= 1048576 || r.end_column >= 16384) return undefined
  return r
}

/** Validate the additive source projection without evaluating worksheet formulas. */
export function decodeNativeConditionalFillPreviewsV1(input: unknown): NativeConditionalFillPreviewV1[] {
  const value = snapshotNativePlainData(input, { maxDepth: 8, maxNodes: 140000 })
  const fail = (): never => { throw new TypeError('Invalid conditional fill preview') }
  const exact = (value: unknown, keys: string[]) => {
    if (!value || typeof value !== 'object' || Array.isArray(value)) return fail()
    const object = value as Record<string, unknown>
    if (Object.keys(object).length !== keys.length || keys.some(key => !Object.hasOwn(object, key))) return fail()
    return object
  }
  if (!Array.isArray(value) || value.length > 64) return fail()
  const ids = new Set<string>(), parts = new Set<string>()
  let count = 0
  return value.map(raw => {
    const status = (raw as Record<string, unknown> | undefined)?.status
    const o = exact(raw, ['sheet_id', 'sheet_part', 'status', 'warnings', ...(status === 'available' ? ['rule', 'cells'] : [])])
    if (typeof o.sheet_id !== 'string' || !/^[1-9][0-9]{0,9}$/.test(o.sheet_id) || Number(o.sheet_id) > 0xffffffff || ids.has(o.sheet_id) || typeof o.sheet_part !== 'string' || !/^[A-Za-z0-9_./-]{1,1024}$/.test(o.sheet_part) || o.sheet_part.split('/').some(p => !p || p === '.' || p === '..') || parts.has(o.sheet_part)) return fail()
    ids.add(o.sheet_id); parts.add(o.sheet_part)
    if (!Array.isArray(o.warnings) || o.warnings.length < 1 || o.warnings.length > 8 || o.warnings.some(w => typeof w !== 'string' || !w || w.length > 4096 || /[\u0000-\u001f\u007f]/.test(w))) return fail()
    const base = { sheet_id: o.sheet_id, sheet_part: o.sheet_part, warnings: o.warnings as string[] }
    if (status === 'unavailable') return { ...base, status }
    if (status !== 'available') return fail()
    const r = exact(o.rule, ['ref', 'operator', 'operand', 'priority', 'stop_if_true', 'dxf_id', 'fill'])
    const bounds = typeof r.ref === 'string' ? range(r.ref) : undefined
    if (!bounds || typeof r.operator !== 'string' || !OPERATORS.includes(r.operator) || typeof r.operand !== 'string' || !INTEGER.test(r.operand) || !Number.isInteger(r.priority) || Number(r.priority) < 1 || Number(r.priority) > 2147483647 || typeof r.stop_if_true !== 'boolean' || !Number.isInteger(r.dxf_id) || Object.is(r.dxf_id, -0) || Number(r.dxf_id) < 0 || Number(r.dxf_id) > 4095 || typeof r.fill !== 'string' || !/^#[0-9A-F]{6}$/.test(r.fill)) return fail()
    const size = (bounds.end_row - bounds.row + 1) * (bounds.end_column - bounds.column + 1)
    if (size > 4096 || (count += size) > 16384 || !Array.isArray(o.cells) || o.cells.length !== size) return fail()
    const rule: NativeConditionalFillRuleV1 = { ref: r.ref as string, operator: r.operator as NativeConditionalOperatorV1, operand: r.operand, priority: Number(r.priority), stop_if_true: r.stop_if_true, dxf_id: Number(r.dxf_id), fill: r.fill }
    const cells = o.cells.map((raw, index) => {
      const c = exact(raw, ['row', 'column', 'lexical', 'cached', 'matches'])
      const width = bounds.end_column - bounds.column + 1
      if (c.row !== bounds.row + Math.floor(index / width) || c.column !== bounds.column + index % width || Object.is(c.row, -0) || Object.is(c.column, -0) || typeof c.lexical !== 'string' || !INTEGER.test(c.lexical) || typeof c.cached !== 'boolean' || typeof c.matches !== 'boolean' || c.matches !== compare(Number(c.lexical), Number(rule.operand), rule.operator)) return fail()
      return { row: Number(c.row), column: Number(c.column), lexical: c.lexical, cached: c.cached, matches: c.matches }
    })
    return { ...base, status, rule, cells }
  })
}

/** Source-bound overlay selection. Every evaluated source value must still join
 * the opened worksheet, including cache ownership. A mismatch refuses the whole
 * overlay; no source styles or formula caches are changed. */
export function selectNativeConditionalFillPreviewV1(workbook: NativeWorkbookV2, sheetId: string, objects: NativeWorkbookObjectsV1): NativeConditionalFillPreviewV1 | undefined {
  if (objects.protocol !== 'injoffice.xlsx.preview-objects' || objects.version !== 1 || !/^sha256:[a-f0-9]{64}$/.test(workbook.source.package_sha256) || objects.package_sha256 !== workbook.source.package_sha256) throw new TypeError('Conditional fill preview does not join the opened source')
  const sheets = workbook.sheets.filter(sheet => sheet.id === sheetId)
  if (sheets.length !== 1) throw new TypeError('Conditional fill preview worksheet identity is unavailable')
  const sheet = sheets[0]!
  const entries = decodeNativeConditionalFillPreviewsV1(objects.conditional_fills ?? [])
  const entry = entries.find(entry => entry.sheet_id === sheet.id || entry.sheet_part === sheet.part_name)
  if (!entry) return undefined
  if (entry.sheet_id !== sheet.id || entry.sheet_part !== sheet.part_name) throw new TypeError('Conditional fill preview worksheet identity mismatch')
  if (entry.status === 'unavailable') return entry
  if (objects.tables.some(table => table.sheet_part === sheet.part_name)) throw new TypeError('Conditional fill preview does not support table combinations')
  const bounds = range(entry.rule.ref)!
  if (sheet.merged_ranges.some(merge => merge.row <= bounds.end_row && merge.end_row >= bounds.row && merge.column <= bounds.end_column && merge.end_column >= bounds.column)) throw new TypeError('Conditional fill preview intersects merged cells')
  const cells = new Map<string, (typeof sheet.cells)[number]>()
  for (const cell of sheet.cells) {
    const identity = range(cell.ref)
    const key = `${cell.row}:${cell.column}`
    if (!identity || identity.row !== cell.row || identity.column !== cell.column || identity.end_row !== cell.row || identity.end_column !== cell.column || cells.has(key)) throw new TypeError('Conditional fill preview cell identity mismatch')
    if (cell.formula && cell.formula.type !== 'normal') throw new TypeError('Conditional fill preview does not support formula groups')
    cells.set(key, cell)
  }
  for (const evidence of entry.cells) {
    const cell = cells.get(`${evidence.row}:${evidence.column}`)
    const value = cell?.formula ? cell.formula.cached : cell?.value
    if (workbook.unsupported.some(item => item.part_name === sheet.part_name && item.cell_ref === cell?.ref)) throw new TypeError('Conditional fill preview source cell has unsupported metadata or markup')
    if (!cell || Boolean(cell.formula) !== evidence.cached || cell.formula && cell.formula.type !== 'normal' || value?.kind !== 'number' || value.storage !== 'number' || value.lexical !== evidence.lexical) throw new TypeError('Conditional fill preview no longer joins every saved source value')
  }
  return entry
}
