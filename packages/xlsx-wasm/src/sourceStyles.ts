/** Separate read-only source projection. Hashes join trusted worker output to
 * source bytes; they are identities, not signatures for arbitrary JSON. */
type Read<T> = (value: unknown) => T
function fail(): never { throw new TypeError('Invalid XLSX source-style preview.') }
function string(max: number, pattern?: RegExp): Read<string> {
  return value => typeof value === 'string' && value.length <= max && wellFormed(value) && (!pattern || pattern.test(value)) ? value : fail()
}
function wellFormed(value: string): boolean {
  for (let i = 0; i < value.length; i++) {
    const unit = value.charCodeAt(i)
    if (unit >= 0xd800 && unit <= 0xdbff) { const next = value.charCodeAt(++i); if (!(next >= 0xdc00 && next <= 0xdfff)) return false }
    else if (unit >= 0xdc00 && unit <= 0xdfff) return false
  }
  return true
}
function number(min: number, max: number, integer = true): Read<number> {
  return value => typeof value === 'number' && Number.isFinite(value) && value >= min && value <= max && (!integer || Number.isSafeInteger(value)) ? value : fail()
}
const boolean: Read<boolean> = value => typeof value === 'boolean' ? value : fail()
function literal<const T extends string | number | boolean>(expected: T): Read<T> { return value => value === expected ? expected : fail() }
function oneOf<const T extends string>(...values: T[]): Read<T> { return value => typeof value === 'string' && values.includes(value as T) ? value as T : fail() }
function array<T>(read: Read<T>, max: number, min = 0): Read<readonly T[]> {
  return value => Array.isArray(value) && value.length >= min && value.length <= max ? Object.freeze(value.map(read)) : fail()
}
function shape<S extends Record<string, Read<unknown>>>(fields: S): Read<{ readonly [K in keyof S]: ReturnType<S[K]> }> {
  return value => {
    if (!value || typeof value !== 'object' || Array.isArray(value)) return fail()
    const record = value as Record<string, unknown>
    if (Object.keys(record).length !== Object.keys(fields).length || Object.keys(record).some(key => !Object.hasOwn(fields, key))) return fail()
    const output: Record<string, unknown> = {}
    for (const key of Object.keys(fields)) output[key] = fields[key]!(record[key])
    return Object.freeze(output) as { readonly [K in keyof S]: ReturnType<S[K]> }
  }
}
const hash = string(71, /^sha256:[a-f0-9]{64}$/)
const rgb = string(7, /^#[A-F0-9]{6}$/)
const fill = string(7, /^(?:#[A-F0-9]{6})?$/)
const id = number(0, 127)
const warnings = array(string(4096), 128)
const borders: Read<Readonly<Partial<Record<'left' | 'right' | 'top' | 'bottom', string>>>> = value => {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return fail()
  const output: Record<string, string> = {}
  for (const [key, color] of Object.entries(value)) {
    if (!['left', 'right', 'top', 'bottom'].includes(key)) return fail()
    output[key] = rgb(color)
  }
  return Object.freeze(output)
}
const readStyle = shape({
  id, parent_id: id, parent_sha256: hash, font_id: id, fill_id: id, border_id: id,
  raw_sha256: hash, font_sha256: hash, fill_sha256: hash, border_sha256: hash,
  number_format_id: number(0, 2147483647), number_format: string(1024),
  font_name: string(128, /^[A-Za-z0-9][A-Za-z0-9 ._-]*$/), font_size_points: number(1, 409, false), font_color: rgb,
  bold: boolean, italic: boolean, fill_color: fill,
  horizontal: oneOf('general', 'left', 'center', 'right'), vertical: oneOf('top', 'center', 'bottom'), wrap: boolean,
  borders, warnings,
})
const readCell = shape({
  ref: string(5, /^[A-Z]{1,2}[1-9][0-9]{0,2}$/), row: number(0, 127), column: number(0, 31), style_id: id,
  kind: oneOf('blank', 'number', 'string', 'boolean', 'error', 'date'), text: string(4096), lexical: string(256), formula: string(2048), cached: boolean,
})
const readMerge = shape({ row: number(0, 127), column: number(0, 31), end_row: number(0, 127), end_column: number(0, 31) })
const readSheet = shape({
  id: string(128), name: string(255), part: string(1024),
  row_heights: array(number(Number.MIN_VALUE, 409, false), 128, 1), column_widths: array(number(Number.MIN_VALUE, 255, false), 32, 1),
  cells: array(readCell, 4096), merges: array(readMerge, 128),
})
const readPreview = shape({
  protocol: literal('injoffice.xlsx.source-style-preview'), version: literal(1), read_only: literal(true), fidelity: literal('approximate'),
  package_sha256: hash, workbook_part: string(1024), styles_part: string(1024), date1904: boolean, strict_error: string(4096), warnings,
  conflicts: array(shape({ style_id: id, parent_id: id, component: oneOf('fill', 'number-format'), apply_flag: oneOf('applyFill', 'applyNumberFormat'), parent_component_id: number(0, 2147483647), direct_component_id: number(0, 2147483647) }), 256, 1),
  styles: array(readStyle, 128, 1), sheets: array(readSheet, 1, 1),
})
export type XlsxSourceStylePreviewV1 = ReturnType<typeof readPreview>
export type XlsxSourceStyleV1 = ReturnType<typeof readStyle>
export type XlsxSourceStyleCellV1 = ReturnType<typeof readCell>

/** Decode only JSON from the dedicated native preview worker. The returned
 * snapshot is recursively frozen and contains no native mutation authority. */
export function decodeXlsxSourceStylePreviewV1(json: string, packageSHA256: string): XlsxSourceStylePreviewV1 {
  if (typeof json !== 'string' || json.length > 4 * 1024 * 1024) return fail()
  hash(packageSHA256)
  const preview = readPreview(JSON.parse(json) as unknown)
  if (preview.package_sha256 !== packageSHA256 || !preview.strict_error || !preview.workbook_part || !preview.styles_part || preview.warnings.length < 2) return fail()
  const styles = new Map(preview.styles.map(style => [style.id, style]))
  if (styles.size !== preview.styles.length || !styles.has(0)) return fail()
  const conflicts = new Set<string>()
  for (const conflict of preview.conflicts) {
    const key = `${conflict.style_id}:${conflict.component}`
    if (conflicts.has(key) || conflict.parent_component_id === conflict.direct_component_id || conflict.apply_flag !== (conflict.component === 'fill' ? 'applyFill' : 'applyNumberFormat')) return fail()
    if (conflict.component === 'fill' && (conflict.parent_component_id > 127 || conflict.direct_component_id > 127)) return fail()
    conflicts.add(key)
    const style = styles.get(conflict.style_id)
    if (style && (style.parent_id !== conflict.parent_id || (conflict.component === 'fill' ? style.fill_id : style.number_format_id) !== conflict.direct_component_id)) return fail()
  }
  const sheet = preview.sheets[0]!
  if (!sheet.id || !sheet.name || !sheet.part) return fail()
  const cells = new Map<string, XlsxSourceStyleCellV1>()
  let totalText = 0
  for (const cell of sheet.cells) {
    const column = cell.column < 26 ? String.fromCharCode(65 + cell.column) : 'A' + String.fromCharCode(65 + cell.column - 26)
    if (cell.ref !== `${column}${cell.row + 1}` || cell.row >= sheet.row_heights.length || cell.column >= sheet.column_widths.length || !styles.has(cell.style_id) || cells.has(cell.ref)) return fail()
    if (cell.cached !== (cell.formula.length > 0)) return fail()
    if (cell.kind === 'boolean' && cell.lexical !== '0' && cell.lexical !== '1') return fail()
    if (cell.kind === 'number' && (!/^[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[Ee][+-]?[0-9]+)?$/.test(cell.lexical) || !Number.isFinite(Number(cell.lexical)))) return fail()
    if (cell.kind === 'blank' && (cell.text !== '' || cell.lexical !== '' || cell.cached)) return fail()
    totalText += cell.text.length + cell.lexical.length + cell.formula.length
    if (totalText > 65536) return fail()
    cells.set(cell.ref, cell)
  }
  const coverage = new Set<string>()
  for (const merge of sheet.merges) {
    if (merge.end_row < merge.row || merge.end_column < merge.column || merge.end_row >= sheet.row_heights.length || merge.end_column >= sheet.column_widths.length || merge.end_row === merge.row && merge.end_column === merge.column) return fail()
    for (let row = merge.row; row <= merge.end_row; row++) for (let col = merge.column; col <= merge.end_column; col++) {
      const key = `${row}:${col}`
      if (coverage.has(key)) return fail()
      coverage.add(key)
    }
    for (const cell of sheet.cells) if (cell.row >= merge.row && cell.row <= merge.end_row && cell.column >= merge.column && cell.column <= merge.end_column && (cell.row !== merge.row || cell.column !== merge.column) && cell.kind !== 'blank') return fail()
  }
  return preview
}
