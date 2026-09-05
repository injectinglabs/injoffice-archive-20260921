import {
  EXCEL_MAX_COLUMNS,
  EXCEL_MAX_ROWS,
} from './mutationProtocol.js'
import { XLSX_NATIVE_RESOURCE_LIMITS, XLSX_NATIVE_UNSUPPORTED_CLASSIFICATIONS } from './nativeContract.generated.js'
import type {
  NativeWorkbookCellV1,
  NativeWorkbookFormulaV1,
  NativeWorkbookMergedRangeV1,
  NativeWorkbookEffectiveStyleV1,
  NativeWorkbookUnsupportedV1,
  NativeWorkbookV1,
  NativeWorkbookValueV1,
} from './nativeContract.generated.js'
import { parseNativeWorkbookJson, utf8ByteLength } from './nativeJson.js'
import { sha256Hex } from './nativeSha256.js'
import { validateNativeWorkbookShape } from './nativeSchemaValidation.js'
import type { NativeWorkbookValidationIssue, ValidateNativeWorkbookResult } from './nativeSchemaValidation.js'

export type DecodeNativeWorkbookResult = ValidateNativeWorkbookResult

export class NativeWorkbookValidationError extends Error {
  readonly issues: ReadonlyArray<NativeWorkbookValidationIssue>
  constructor(issues: ReadonlyArray<NativeWorkbookValidationIssue>) {
    super(issues.slice(0, 8).map((item) => `${item.path}: ${item.message}`).join('; '))
    this.name = 'NativeWorkbookValidationError'
    this.issues = issues
  }
}

export function decodeNativeWorkbookV1(source: string): DecodeNativeWorkbookResult {
  let input: unknown
  try { input = parseNativeWorkbookJson(source) }
  catch (error) { return { ok: false, issues: [{ code: 'INVALID_JSON', path: '/', message: error instanceof Error ? error.message : 'invalid JSON' }] } }
  return validateNativeWorkbookV1(input)
}

export function validateNativeWorkbookV1(input: unknown): ValidateNativeWorkbookResult {
  const shaped = validateNativeWorkbookShape(input)
  if (!shaped.ok) return shaped
  const issues: NativeWorkbookValidationIssue[] = []
  validateSemantics(shaped.value, issues)
  return issues.length === 0 ? shaped : { ok: false, issues }
}

export function assertNativeWorkbookV1(input: unknown): asserts input is NativeWorkbookV1 {
  const result = validateNativeWorkbookV1(input)
  if (!result.ok) throw new NativeWorkbookValidationError(result.issues)
}

function validateSemantics(workbook: NativeWorkbookV1, issues: NativeWorkbookValidationIssue[]): void {
  if (workbook.revision.slice(4) !== workbook.source.package_sha256.slice(7)) add(issues, 'INVALID_REVISION', '/revision', 'revision must fingerprint the exact source package bytes')
  if (!canonicalOpcPartKey(workbook.source.workbook_part)) add(issues, 'INVALID_PART', '/source/workbook_part', 'must be a canonical OPC part URI')

  const sheetIDs = new Set<string>(), sheetParts = new Set<string>(), sheetNames = new Set<string>()
  let totalRows = 0, totalColumns = 0, totalCells = 0, totalMergedRanges = 0
  for (let sheetIndex = 0; sheetIndex < workbook.sheets.length; sheetIndex++) {
    const sheet = workbook.sheets[sheetIndex]
    const path = `/sheets/${sheetIndex}`
    if (sheet.order !== sheetIndex) add(issues, 'INVALID_ORDER', `${path}/order`, 'sheet order must match array order')
    if (!isCanonicalSheetID(sheet.id) || sheetIDs.has(sheet.id)) add(issues, 'INVALID_ID', `${path}/id`, 'sheet id must be a unique canonical uint32')
    sheetIDs.add(sheet.id)
    const sheetNameKey = nativeSheetNameCaseKey(sheet.name)
    if (/[\[\]:*?\/\\]/.test(sheet.name) || [...sheet.name].some((item) => item.codePointAt(0)! < 0x20) || sheet.name.startsWith("'") || sheet.name.endsWith("'")) add(issues, 'INVALID_VALUE', `${path}/name`, 'sheet name contains an invalid Excel character or boundary apostrophe')
    if (sheetNames.has(sheetNameKey)) add(issues, 'DUPLICATE_ID', `${path}/name`, 'sheet names must be case-insensitively unique')
    sheetNames.add(sheetNameKey)
    const partKey = canonicalOpcPartKey(sheet.part_name)
    if (!partKey || sheetParts.has(partKey)) add(issues, 'DUPLICATE_PART', `${path}/part_name`, 'worksheet part is invalid or case/escape-equivalent to another sheet')
    else sheetParts.add(partKey)
    if (sheet.editable === (sheet.refusal_code !== undefined)) add(issues, 'INVALID_UNION', `${path}/refusal_code`, 'editable/refusal_code state is inconsistent')
    if (sheet.sheet_format !== undefined) {
      if (Object.is(sheet.sheet_format.default_row_height_points, -0)) add(issues, 'INVALID_NUMBER', `${path}/sheet_format/default_row_height_points`, 'negative zero is not canonical')
      if (sheet.sheet_format.default_column_width !== undefined && Object.is(sheet.sheet_format.default_column_width, -0)) add(issues, 'INVALID_NUMBER', `${path}/sheet_format/default_column_width`, 'negative zero is not canonical')
    }
    totalRows += sheet.rows.length; totalColumns += sheet.columns.length; totalCells += sheet.cells.length; totalMergedRanges += sheet.merged_ranges.length

    let lastRow = -1
    for (let index = 0; index < sheet.rows.length; index++) {
      const row = sheet.rows[index]
      if (row.row <= lastRow) add(issues, 'INVALID_ORDER', `${path}/rows/${index}/row`, 'row dimensions must be unique and ordered')
      lastRow = row.row
      if (row.custom_height && row.height_points === undefined) add(issues, 'INVALID_UNION', `${path}/rows/${index}`, 'custom_height requires height_points')
      if (row.style_id !== undefined && row.style_id >= workbook.styles.length) add(issues, 'INVALID_REFERENCE', `${path}/rows/${index}/style_id`, 'style id is outside styles')
    }
    let lastColumn = -1
    for (let index = 0; index < sheet.columns.length; index++) {
      const column = sheet.columns[index]
      if (column.column <= lastColumn || column.end_column < column.column) add(issues, 'INVALID_ORDER', `${path}/columns/${index}`, 'column ranges must be non-overlapping and ordered')
      lastColumn = column.end_column
      if (column.custom_width && column.width === undefined) add(issues, 'INVALID_UNION', `${path}/columns/${index}`, 'custom_width requires width')
      if (column.style_id !== undefined && column.style_id >= workbook.styles.length) add(issues, 'INVALID_REFERENCE', `${path}/columns/${index}/style_id`, 'style id is outside styles')
    }
    let lastCell = -1
    for (let index = 0; index < sheet.cells.length; index++) {
      const cell = sheet.cells[index]
      const position = cell.row * EXCEL_MAX_COLUMNS + cell.column
      if (position <= lastCell) add(issues, 'INVALID_ORDER', `${path}/cells/${index}`, 'cells must be unique and row-major ordered')
      lastCell = position
      if (cell.ref !== cellReference(cell.row, cell.column)) add(issues, 'INVALID_REFERENCE', `${path}/cells/${index}/ref`, 'cell reference does not match coordinates')
      if (cell.style_id >= workbook.styles.length) add(issues, 'INVALID_REFERENCE', `${path}/cells/${index}/style_id`, 'style id is outside styles')
      validateCell(cell, `${path}/cells/${index}`, sheet.editable, issues)
    }
    const mergedRanges: Range[] = []
    let lastMergedPosition = -1
    for (let index = 0; index < sheet.merged_ranges.length; index++) {
      const merged = sheet.merged_ranges[index]
      const mergedPath = `${path}/merged_ranges/${index}`
      const position = merged.row * EXCEL_MAX_COLUMNS + merged.column
      if (position <= lastMergedPosition) add(issues, 'INVALID_ORDER', mergedPath, 'merged ranges must be unique and row-major ordered by their top-left cell')
      lastMergedPosition = position
      const range = validateMergedRange(merged, mergedPath, issues)
      if (range) mergedRanges.push(range)
    }
    if (hasOverlappingRanges(mergedRanges)) add(issues, 'INVALID_UNION', `${path}/merged_ranges`, 'merged ranges must not overlap')
  }
  if (totalRows > XLSX_NATIVE_RESOURCE_LIMITS.maxRows || totalColumns > XLSX_NATIVE_RESOURCE_LIMITS.maxCells || totalCells > XLSX_NATIVE_RESOURCE_LIMITS.maxCells || totalMergedRanges > XLSX_NATIVE_RESOURCE_LIMITS.maxMergedRanges) {
    add(issues, 'LIMIT_EXCEEDED', '/sheets', 'cumulative row, column, cell, or merged-range inventory exceeds its resource bound')
  }
  validateStyles(workbook, issues)
  validateCapabilities(workbook, issues)
  validatePassthrough(workbook, issues)
  validateUnsupportedAuthority(workbook, issues)
  if (nativeContractTextBytes(workbook) > XLSX_NATIVE_RESOURCE_LIMITS.maxTextBytes) add(issues, 'LIMIT_EXCEEDED', '/', 'contract strings exceed the cumulative byte bound')
}

function validateMergedRange(merged: NativeWorkbookMergedRangeV1, path: string, issues: NativeWorkbookValidationIssue[]): Range | undefined {
  if (merged.editable !== false) add(issues, 'INVALID_VALUE', `${path}/editable`, 'merged ranges are structurally mutation-refused in native v1')
  if (merged.end_row < merged.row || merged.end_column < merged.column || (merged.end_row === merged.row && merged.end_column === merged.column)) {
    add(issues, 'INVALID_REFERENCE', path, 'merged range must span at least two cells in forward grid order')
    return undefined
  }
  const range = parseRange(merged.ref)
  if (!range || merged.ref !== range.canonical || range.minRow !== merged.row || range.minColumn !== merged.column || range.maxRow !== merged.end_row || range.maxColumn !== merged.end_column) {
    add(issues, 'INVALID_REFERENCE', `${path}/ref`, 'merged range ref must be canonical bounded A1 and exactly match its coordinates')
    return undefined
  }
  return range
}

function validateCell(cell: NativeWorkbookCellV1, path: string, sheetEditable: boolean, issues: NativeWorkbookValidationIssue[]): void {
  if (cell.value !== undefined && cell.formula !== undefined) add(issues, 'INVALID_UNION', path, 'cell cannot contain both a literal value and a formula')
  const storage = cell.ooxml_type ?? 'n'
  if (cell.value) validateValue(cell.value, storage, `${path}/value`, issues)
  if (!cell.formula) {
    if (!cell.value && (storage === 's' || storage === 'inlineStr')) add(issues, 'REQUIRED', `${path}/value`, 'shared and inline string cells require a value')
    return
  }
  const formula = cell.formula
  if (formula.cached) validateValue(formula.cached, storage, `${path}/formula/cached`, issues)
  if (storage === 's' || storage === 'inlineStr') add(issues, 'INVALID_UNION', `${path}/ooxml_type`, 'formula cells cannot use shared or inline string storage')
  if (formula.type !== 'normal' && (cell.editable || sheetEditable)) add(issues, 'INVALID_VALUE', `${path}/editable`, 'group formula cells and sheets must be mutation-refused')
  if ((formula.type === 'normal' && (formula.ref !== undefined || formula.shared_index !== undefined)) ||
      ((formula.type === 'array' || formula.type === 'dataTable') && formula.shared_index !== undefined)) add(issues, 'INVALID_UNION', `${path}/formula`, 'formula attributes do not match formula type')
  if (formula.type === 'shared') {
    if (formula.shared_index === undefined) add(issues, 'REQUIRED', `${path}/formula/shared_index`, 'shared formulas require shared_index')
    if (formula.text === '' && formula.ref !== undefined) add(issues, 'INVALID_UNION', `${path}/formula/ref`, 'shared followers cannot carry a range')
    if (formula.text !== '' && formula.ref === undefined) add(issues, 'REQUIRED', `${path}/formula/ref`, 'shared masters require a range')
  }
  if ((formula.type === 'array' || formula.type === 'dataTable') && formula.ref === undefined) add(issues, 'REQUIRED', `${path}/formula/ref`, 'formula group requires a range')
  if ((formula.type === 'normal' || formula.type === 'array') && formula.text === '') add(issues, 'REQUIRED', `${path}/formula/text`, 'formula text is required')
  if (formula.ref !== undefined) {
    const range = parseRange(formula.ref)
    if (!range || formula.ref !== range.canonical || cell.row < range.minRow || cell.row > range.maxRow || cell.column < range.minColumn || cell.column > range.maxColumn) add(issues, 'INVALID_REFERENCE', `${path}/formula/ref`, 'formula cell must be inside a canonical bounded group range')
  }
}

function validateValue(value: NativeWorkbookValueV1, ooxmlType: string, path: string, issues: NativeWorkbookValidationIssue[]): void {
  const mapping = {
    number: ['number', 'n', true, false, false], boolean: ['boolean', 'b', true, false, false], error: ['error', 'e', true, false, false], date: ['date', 'd', true, false, false],
    'formula-string': ['string', 'str', true, true, false], shared: ['string', 's', true, true, true], inline: ['string', 'inlineStr', false, true, true],
  } as const
  const expected = mapping[value.storage]
  if (value.kind !== expected[0] || ooxmlType !== expected[1]) add(issues, 'INVALID_UNION', path, 'value kind/storage/OOXML type disagree')
  if ((value.lexical !== undefined) !== expected[2] || (value.text !== undefined) !== expected[3] || (value.rich && !expected[4])) add(issues, 'INVALID_UNION', path, 'value lexical/text/rich fields do not match storage')
  if (value.lexical !== undefined) {
    if (value.storage === 'number' && (!finiteNumberLexical(value.lexical))) add(issues, 'INVALID_VALUE', `${path}/lexical`, 'invalid finite OOXML numeric lexical')
    if (value.storage === 'boolean' && !['0', '1', 'false', 'true'].includes(value.lexical)) add(issues, 'INVALID_VALUE', `${path}/lexical`, 'invalid OOXML boolean lexical')
    if (value.storage === 'error' && ([...value.lexical].some((item) => item.codePointAt(0)! < 0x20 && item !== '\t'))) add(issues, 'INVALID_VALUE', `${path}/lexical`, 'invalid OOXML error lexical')
    if (value.storage === 'date' && !validIsoDateTime(value.lexical)) add(issues, 'INVALID_VALUE', `${path}/lexical`, 'date cells require an ISO-8601 date or dateTime lexical')
    if (value.storage === 'shared' && !validUint32Lexical(value.lexical)) add(issues, 'INVALID_VALUE', `${path}/lexical`, 'shared-string lexical must be a uint32 decimal index')
    if (value.storage === 'formula-string' && value.text !== undefined && decodeSpreadsheetString(value.lexical) !== value.text) add(issues, 'INVALID_UNION', path, 'formula-string lexical and decoded text disagree')
  }
}

function validateStyles(workbook: NativeWorkbookV1, issues: NativeWorkbookValidationIssue[]): void {
  for (let index = 0; index < workbook.styles.length; index++) {
    const style = workbook.styles[index]
    const path = `/styles/${index}`
    if (style.id !== index) add(issues, 'INVALID_ORDER', `${path}/id`, 'style ids must be dense and ordered')
    if (style.raw_projection_sha256 !== nativeWorkbookStyleRawProjectionSha256V1(style.effective)) add(issues, 'INVALID_SHA256', `${path}/raw_projection_sha256`, 'style raw projection digest does not match the canonical effective-style bytes')
    const unsupported = style.effective.unsupported
    for (let item = 1; item < unsupported.length; item++) if (unsupported[item] <= unsupported[item - 1]) add(issues, 'INVALID_ORDER', `${path}/effective/unsupported/${item}`, 'unsupported style codes must be unique and sorted')
    if ((style.effective.projection === 'full') !== (unsupported.length === 0)) add(issues, 'INVALID_UNION', `${path}/effective`, 'projection must match unsupported components')
    const borderUnsupported = unsupported.includes('border')
    const fillUnsupported = unsupported.includes('fill')
    if ((style.effective.fill === undefined) !== fillUnsupported) add(issues, 'INVALID_UNION', `${path}/effective/fill`, 'fill provenance must be present exactly when fill is supported')
    const fill = style.effective.fill
    if (fill?.origin === 'implicit-default' && (fill.fill_id !== undefined || fill.record_sha256 !== undefined || fill.color !== undefined || style.effective.fill_color !== undefined)) {
      add(issues, 'INVALID_UNION', `${path}/effective/fill`, 'implicit default fill cannot carry styles-table provenance or color')
    }
    if (fill?.origin === 'styles-record' && (fill.fill_id === undefined || fill.record_sha256 === undefined)) {
      add(issues, 'REQUIRED', `${path}/effective/fill`, 'styles-record fill requires its exact fill id and record digest')
    }
    if ((fill?.color === undefined) !== (style.effective.fill_color === undefined) || (fill?.color !== undefined && fill.color !== style.effective.fill_color)) {
      add(issues, 'INVALID_UNION', `${path}/effective/fill_color`, 'effective fill color must exactly match fill provenance')
    }
    if ((style.effective.border === undefined) !== borderUnsupported) add(issues, 'INVALID_UNION', `${path}/effective/border`, 'border projection must be present exactly when border is supported')
    const border = style.effective.border
    if (border?.origin === 'implicit-default' && (border.border_id !== undefined || border.record_sha256 !== undefined || border.left !== undefined || border.right !== undefined || border.top !== undefined || border.bottom !== undefined)) {
      add(issues, 'INVALID_UNION', `${path}/effective/border`, 'implicit default border cannot carry styles-table provenance or sides')
    }
    if (border?.origin === 'styles-record' && (border.border_id === undefined || border.record_sha256 === undefined)) {
      add(issues, 'REQUIRED', `${path}/effective/border`, 'styles-record border requires its exact border id and record digest')
    }
  }
}

/** Canonical compact wire bytes shared with Go's NativeWorkbookEffectiveStyleV1 field order. */
export function nativeWorkbookStyleRawProjectionSha256V1(style: NativeWorkbookEffectiveStyleV1): `sha256:${string}` {
  const canonical = {
    ...(style.number_format === undefined ? {} : { number_format: style.number_format }),
    ...(style.font_name === undefined ? {} : { font_name: style.font_name }),
    ...(style.font_size_points === undefined ? {} : { font_size_points: style.font_size_points }),
    ...(style.bold === undefined ? {} : { bold: style.bold }),
    ...(style.italic === undefined ? {} : { italic: style.italic }),
    ...(style.font_color === undefined ? {} : { font_color: style.font_color }),
    ...(style.fill_color === undefined ? {} : { fill_color: style.fill_color }),
    ...(style.fill === undefined ? {} : { fill: canonicalFill(style.fill) }),
    ...(style.border === undefined ? {} : { border: canonicalBorder(style.border) }),
    ...(style.horizontal_alignment === undefined ? {} : { horizontal_alignment: style.horizontal_alignment }),
    ...(style.vertical_alignment === undefined ? {} : { vertical_alignment: style.vertical_alignment }),
    ...(style.wrap_text === undefined ? {} : { wrap_text: style.wrap_text }),
    projection: style.projection,
    unsupported: [...style.unsupported],
  }
  // encoding/json deliberately escapes these two ECMAScript line separators
  // even with SetEscapeHTML(false); mirror its compact wire bytes exactly.
  const json = JSON.stringify(canonical).replaceAll('\u2028', '\\u2028').replaceAll('\u2029', '\\u2029')
  return `sha256:${sha256Hex(json)}`
}

function canonicalFill(fill: NonNullable<NativeWorkbookEffectiveStyleV1['fill']>): object {
  return {
    origin: fill.origin,
    ...(fill.fill_id === undefined ? {} : { fill_id: fill.fill_id }),
    ...(fill.record_sha256 === undefined ? {} : { record_sha256: fill.record_sha256 }),
    ...(fill.color === undefined ? {} : { color: fill.color }),
  }
}

function canonicalBorder(border: NonNullable<NativeWorkbookEffectiveStyleV1['border']>): object {
  return {
    origin: border.origin,
    ...(border.border_id === undefined ? {} : { border_id: border.border_id }),
    ...(border.record_sha256 === undefined ? {} : { record_sha256: border.record_sha256 }),
    ...(border.left === undefined ? {} : { left: { style: border.left.style, color: border.left.color } }),
    ...(border.right === undefined ? {} : { right: { style: border.right.style, color: border.right.color } }),
    ...(border.top === undefined ? {} : { top: { style: border.top.style, color: border.top.color } }),
    ...(border.bottom === undefined ? {} : { bottom: { style: border.bottom.style, color: border.bottom.color } }),
  }
}

function validateCapabilities(workbook: NativeWorkbookV1, issues: NativeWorkbookValidationIssue[]): void {
  const expected = [['native-ooxml-parse', 'read-only'], ['native-v1-mutations', 'partial'], ['unsupported-content', 'preserve-exact']]
  workbook.capabilities.forEach((item, index) => {
    if (item.name !== expected[index]?.[0] || item.level !== expected[index]?.[1]) add(issues, 'INVALID_VALUE', `/capabilities/${index}`, 'capability name, level, or order is not canonical for v1')
  })
}

function validatePassthrough(workbook: NativeWorkbookV1, issues: NativeWorkbookValidationIssue[]): void {
  const seen = new Set<string>()
  let prior = ''
  workbook.passthrough_parts.forEach((part, index) => {
    const path = `/passthrough_parts/${index}`
    const key = canonicalOpcPartKey(part.part_name)
    if (!key || seen.has(key)) add(issues, 'DUPLICATE_PART', `${path}/part_name`, 'invalid or duplicate case/escape-equivalent passthrough part')
    if (key) seen.add(key)
    const order = asciiLower(part.part_name)
    if (index > 0 && order <= prior) add(issues, 'INVALID_ORDER', `${path}/part_name`, 'passthrough parts must be uniquely ASCII-case-sorted')
    prior = order
    if (part.content_type.trim() !== part.content_type || /[\r\n\0]/.test(part.content_type)) add(issues, 'INVALID_VALUE', `${path}/content_type`, 'invalid content type')
  })
}

type UnsupportedClass = { readonly capability: string; readonly scope: 'workbook' | 'sheet' | 'style'; readonly impact: 'none' | 'cell' | 'range' | 'sheet'; readonly refusalCode?: string }
const unsupportedClasses: Readonly<Record<string, UnsupportedClass>> = XLSX_NATIVE_UNSUPPORTED_CLASSIFICATIONS
const effectiveStyleUnsupportedCodes: Readonly<Record<string, string>> = Object.freeze({
  'alignment-extended': 'STYLE_ALIGNMENT_EXTENDED',
  border: 'STYLE_BORDER',
  fill: 'STYLE_FILL',
  'font-color': 'STYLE_FONT_COLOR',
  'horizontal-alignment': 'STYLE_HORIZONTAL_ALIGNMENT',
  'number-format': 'STYLE_NUMBER_FORMAT',
  'vertical-alignment': 'STYLE_VERTICAL_ALIGNMENT',
})

function validateUnsupportedAuthority(workbook: NativeWorkbookV1, issues: NativeWorkbookValidationIssue[]): void {
  const sheets = new Map(workbook.sheets.map((sheet) => [sheet.id, sheet]))
  const cells = new Map(workbook.sheets.map((sheet) => [sheet.id, new Map(sheet.cells.map((cell) => [cell.ref, cell]))]))
  const cellReasons = new Map(workbook.sheets.map((sheet) => [sheet.id, new Set<string>()]))
  const cellCodes = new Map(workbook.sheets.map((sheet) => [sheet.id, new Map<string, Set<string>>()]))
  const sheetReasons = new Map(workbook.sheets.map((sheet) => [sheet.id, new Set<string>()]))
  const ranges = new Map(workbook.sheets.map((sheet) => [sheet.id, [] as Range[]]))
  const observedRanges = new Map(workbook.sheets.map((sheet) => [sheet.id, new Set<string>()]))
  const mergedRanges = new Map(workbook.sheets.map((sheet) => [sheet.id, sheet.merged_ranges.map((item) => parseRange(item.ref)).filter((item): item is Range => item !== undefined)]))
  const mergedSourceSheets = new Set<string>()
  const observedStyleCodes = workbook.styles.map(() => new Set<string>())
  const seenIDs = new Set<string>(), seenLocations = new Set<string>()

  workbook.unsupported.forEach((item, index) => {
    const path = `/unsupported/${index}`
    const classification = unsupportedClasses[item.code]
    if (!classification || classification.capability !== item.capability) add(issues, 'INVALID_VALUE', path, 'unsupported code and capability are not a canonical v1 pair')
    const sheetID = item.scope_id.startsWith('sheet:') ? item.scope_id.slice(6) : undefined
    const styleLexical = item.scope_id.startsWith('style:') ? item.scope_id.slice(6) : undefined
    const styleID = styleLexical !== undefined && /^(?:0|[1-9][0-9]*)$/.test(styleLexical) ? Number(styleLexical) : undefined
    const scope = sheetID !== undefined ? 'sheet' : styleLexical !== undefined ? 'style' : item.scope_id === 'workbook' ? 'workbook' : undefined
    if (!classification || scope !== classification.scope) add(issues, 'INVALID_REFERENCE', `${path}/scope_id`, 'unsupported code is in the wrong scope')
    const sheet = sheetID === undefined ? undefined : sheets.get(sheetID)
    if (sheetID !== undefined && !sheet) add(issues, 'INVALID_REFERENCE', `${path}/scope_id`, 'sheet scope does not exist')
    if (styleLexical !== undefined && (styleID === undefined || !Number.isSafeInteger(styleID) || styleID < 0 || styleID >= workbook.styles.length)) add(issues, 'INVALID_REFERENCE', `${path}/scope_id`, 'style scope does not exist')
    if (classification?.scope === 'style' && styleID !== undefined && Number.isSafeInteger(styleID) && styleID >= 0 && styleID < workbook.styles.length) observedStyleCodes[styleID]!.add(item.code)
    if (item.part_name === undefined || !canonicalOpcPartKey(item.part_name)) add(issues, 'REQUIRED', `${path}/part_name`, 'unsupported content must identify a canonical source-authoritative part')
    if (sheet !== undefined && item.part_name !== undefined && classification?.scope === 'sheet' && item.code !== 'SHEET_DECLARATION_ATTRIBUTES' && item.part_name !== sheet.part_name) add(issues, 'INVALID_REFERENCE', `${path}/part_name`, 'sheet unsupported content must use the exact worksheet part spelling')
    if (item.code === 'SHEET_DECLARATION_ATTRIBUTES' && item.part_name !== workbook.source.workbook_part) add(issues, 'INVALID_REFERENCE', `${path}/part_name`, 'sheet declaration must identify the exact workbook part')
    if (item.cell_ref !== undefined && item.range_ref !== undefined) add(issues, 'INVALID_UNION', path, 'unsupported location cannot contain both cell_ref and range_ref')
    const hasCell = item.cell_ref !== undefined, hasRange = item.range_ref !== undefined
    if (classification?.impact === 'cell' && (!hasCell || hasRange)) add(issues, 'REQUIRED', `${path}/cell_ref`, 'cell-impact code requires exactly cell_ref')
    if (classification?.impact === 'range' && (!hasRange || hasCell)) add(issues, 'REQUIRED', `${path}/range_ref`, 'range-impact code requires exactly range_ref')
    if (classification && (classification.impact === 'none' || classification.impact === 'sheet') && (hasCell || hasRange)) add(issues, 'INVALID_UNION', path, 'this unsupported code cannot carry a cell/range location')

    if (classification?.impact === 'cell' && sheetID !== undefined && item.cell_ref !== undefined) {
      const cell = cells.get(sheetID)?.get(item.cell_ref)
      if (!cell) add(issues, 'INVALID_REFERENCE', `${path}/cell_ref`, 'cell-scoped source is absent from modeled cells')
      else {
        if (cell.editable) add(issues, 'INVALID_VALUE', `${path}/cell_ref`, 'cell-scoped source requires mutation refusal')
        cellReasons.get(sheetID)!.add(item.cell_ref)
        const codes = cellCodes.get(sheetID)!
        if (!codes.has(item.cell_ref)) codes.set(item.cell_ref, new Set())
        codes.get(item.cell_ref)!.add(item.code)
        validateUnsupportedCell(item, cell, path, issues)
      }
    } else if (classification?.impact === 'range' && sheetID !== undefined && item.range_ref !== undefined) {
      const range = parseRange(item.range_ref)
      if (!range || item.range_ref !== range.canonical) add(issues, 'INVALID_REFERENCE', `${path}/range_ref`, 'range_ref must be canonical bounded A1')
      else ranges.get(sheetID)!.push(range)
      observedRanges.get(sheetID)!.add(item.range_ref)
      sheetReasons.get(sheetID)!.add('FORMULA_GROUPS')
    } else if (classification?.impact === 'sheet' && sheetID) sheetReasons.get(sheetID)!.add(classification.refusalCode!)
    if (item.code === 'MERGED_CELLS' && sheetID !== undefined) mergedSourceSheets.add(sheetID)

    const location = [item.code, item.capability, item.scope_id, item.part_name ?? '', item.cell_ref ?? '', item.range_ref ?? ''].join('\0')
    if (seenLocations.has(location)) add(issues, 'DUPLICATE_ID', path, 'duplicate unsupported location')
    seenLocations.add(location)
    if (seenIDs.has(item.id) || item.id !== `unsupported:${sha256Hex(location)}`) add(issues, 'INVALID_ID', `${path}/id`, 'unsupported id must match its unique canonical source location')
    seenIDs.add(item.id)
  })

  workbook.styles.forEach((style, styleID) => {
    const expected = new Set(style.effective.unsupported.map((code) => effectiveStyleUnsupportedCodes[code]).filter((code): code is string => code !== undefined))
    const observed = observedStyleCodes[styleID]!
    for (const code of expected) if (!observed.has(code)) add(issues, 'REQUIRED', `/styles/${styleID}/effective/unsupported`, `style projection lacks its canonical ${code} source diagnostic`)
    for (const code of observed) if (!expected.has(code)) add(issues, 'INVALID_REFERENCE', `/unsupported`, `${code} contradicts the effective style projection for style:${styleID}`)
  })

  for (const sheet of workbook.sheets) {
    const sheetRanges = ranges.get(sheet.id)!
    if (hasOverlappingRanges(sheetRanges)) add(issues, 'INVALID_UNION', `/sheets/${sheet.order}`, 'formula group ranges must not overlap')
    const covered = coveredCellRefs(sheet.cells, sheetRanges)
    const mergeCovered = coveredCellRefs(sheet.cells, mergedRanges.get(sheet.id)!)
    const mergeAnchors = new Set(sheet.merged_ranges.map((range) => cellReference(range.row, range.column)))
    if ((sheet.merged_ranges.length > 0) !== mergedSourceSheets.has(sheet.id)) add(issues, 'INVALID_REFERENCE', `/sheets/${sheet.order}/merged_ranges`, 'modeled merged ranges and their exact MERGED_CELLS source diagnostic must be present together')
    const expectedRanges = new Set<string>()
    for (const cell of sheet.cells) {
      const expectedCodes = new Set<string>()
      if (cell.value?.rich) expectedCodes.add('RICH_CELL_STRING')
      if (cell.formula?.type !== undefined && cell.formula.type !== 'normal') {
        expectedCodes.add(`FORMULA_${cell.formula.type.toUpperCase()}`)
        if (cell.formula.ref !== undefined) expectedRanges.add(cell.formula.ref)
      }
      if (covered.has(cell.ref) && cell.editable) add(issues, 'INVALID_VALUE', `/sheets/${sheet.order}/cells/${cell.ref}/editable`, 'range refusal requires covered cells to be mutation-refused')
      if (mergeCovered.has(cell.ref) && cell.editable) add(issues, 'INVALID_VALUE', `/sheets/${sheet.order}/cells/${cell.ref}/editable`, 'modeled cells covered by a merged range must be mutation-refused')
      if (mergeCovered.has(cell.ref) && !mergeAnchors.has(cell.ref) && (cell.value !== undefined || cell.formula !== undefined)) add(issues, 'INVALID_UNION', `/sheets/${sheet.order}/cells/${cell.ref}`, 'covered non-anchor merged cell must be blank; a value or formula is ambiguous')
      if (!cell.editable && !cellReasons.get(sheet.id)!.has(cell.ref) && !covered.has(cell.ref) && !mergeCovered.has(cell.ref)) add(issues, 'REQUIRED', `/sheets/${sheet.order}/cells/${cell.ref}/editable`, 'non-editable cell requires authoritative unsupported or merged-range source')
      for (const code of expectedCodes) if (!cellCodes.get(sheet.id)!.get(cell.ref)?.has(code)) add(issues, 'REQUIRED', `/sheets/${sheet.order}/cells/${cell.ref}`, 'modeled rich/formula refusal lacks its canonical unsupported source')
    }
    for (const range of expectedRanges) if (!observedRanges.get(sheet.id)!.has(range)) add(issues, 'REQUIRED', `/sheets/${sheet.order}`, 'modeled formula group lacks its exact range refusal')
    for (const range of observedRanges.get(sheet.id)!) if (!expectedRanges.has(range)) add(issues, 'INVALID_REFERENCE', `/sheets/${sheet.order}`, 'formula range refusal does not match modeled formula group')
    const refusal = expectedSheetRefusal(sheetReasons.get(sheet.id)!)
    if (refusal === undefined ? (!sheet.editable || sheet.refusal_code !== undefined) : (sheet.editable || sheet.refusal_code !== refusal)) add(issues, 'INVALID_VALUE', `/sheets/${sheet.order}/refusal_code`, 'sheet refusal does not match authoritative unsupported source')
  }
}

function validateUnsupportedCell(item: NativeWorkbookUnsupportedV1, cell: NativeWorkbookCellV1, path: string, issues: NativeWorkbookValidationIssue[]): void {
  if (item.code === 'RICH_CELL_STRING' && !cell.value?.rich) add(issues, 'INVALID_REFERENCE', `${path}/cell_ref`, 'rich refusal does not reference a rich value')
  if (item.code === 'FORMULA_ATTRIBUTES' && !cell.formula) add(issues, 'INVALID_REFERENCE', `${path}/cell_ref`, 'formula refusal does not reference a formula')
  const expected: Record<string, NativeWorkbookFormulaV1['type']> = { FORMULA_SHARED: 'shared', FORMULA_ARRAY: 'array', FORMULA_DATATABLE: 'dataTable' }
  if (expected[item.code] && cell.formula?.type !== expected[item.code]) add(issues, 'INVALID_REFERENCE', `${path}/cell_ref`, 'formula refusal type differs from modeled formula')
}

type Range = { minRow: number; minColumn: number; maxRow: number; maxColumn: number; ref: string; canonical: string }
function parseRange(value: string): Range | undefined {
  const parts = value.replace(/\$/g, '').split(':')
  if (parts.length < 1 || parts.length > 2) return undefined
  const first = parseCellRef(parts[0]), last = parseCellRef(parts[1] ?? parts[0])
  if (!first || !last || first.row > last.row || first.column > last.column) return undefined
  const firstRef = cellReference(first.row, first.column), lastRef = cellReference(last.row, last.column)
  return { minRow: first.row, minColumn: first.column, maxRow: last.row, maxColumn: last.column, ref: value, canonical: firstRef === lastRef ? firstRef : `${firstRef}:${lastRef}` }
}
function parseCellRef(value: string): { row: number; column: number } | undefined {
  const match = /^([A-Za-z]+)([0-9]+)$/.exec(value)
  if (!match) return undefined
  let column = 0
  for (const character of match[1].toUpperCase()) column = column * 26 + character.charCodeAt(0) - 64
  const row = Number(match[2])
  if (!Number.isSafeInteger(row) || row < 1 || row > EXCEL_MAX_ROWS || column < 1 || column > EXCEL_MAX_COLUMNS) return undefined
  return { row: row - 1, column: column - 1 }
}
function cellReference(row: number, column: number): string {
  let value = column + 1, name = ''
  while (value > 0) { value--; name = String.fromCharCode(65 + value % 26) + name; value = Math.floor(value / 26) }
  return `${name}${row + 1}`
}

function coveredCellRefs(cells: ReadonlyArray<NativeWorkbookCellV1>, ranges: ReadonlyArray<Range>): Set<string> {
  const events = ranges.flatMap((range) => [{ row: range.minRow, start: range.minColumn, end: range.maxColumn, delta: 1 }, { row: range.maxRow + 1, start: range.minColumn, end: range.maxColumn, delta: -1 }]).sort((a, b) => a.row - b.row || a.delta - b.delta)
  const tree = new Int32Array(EXCEL_MAX_COLUMNS + 2), covered = new Set<string>()
  const update = (index: number, delta: number) => { for (let item = index + 1; item < tree.length; item += item & -item) tree[item] += delta }
  const query = (index: number) => { let total = 0; for (let item = index + 1; item > 0; item -= item & -item) total += tree[item]; return total }
  let eventIndex = 0
  for (const cell of cells) {
    while (eventIndex < events.length && events[eventIndex].row <= cell.row) { const event = events[eventIndex++]; update(event.start, event.delta); update(event.end + 1, -event.delta) }
    if (query(cell.column) > 0) covered.add(cell.ref)
  }
  return covered
}

function hasOverlappingRanges(ranges: ReadonlyArray<Range>): boolean {
  if (ranges.length < 2) return false
  const events = ranges.flatMap((range) => [
    { row: range.minRow, start: range.minColumn, end: range.maxColumn, delta: 1 },
    { row: range.maxRow + 1, start: range.minColumn, end: range.maxColumn, delta: -1 },
  ]).sort((left, right) => left.row - right.row || left.delta - right.delta || left.start - right.start || left.end - right.end)
  const maximum = new Int32Array(EXCEL_MAX_COLUMNS * 4), lazy = new Int32Array(EXCEL_MAX_COLUMNS * 4)
  const addRange = (node: number, left: number, right: number, start: number, end: number, delta: number): void => {
    if (start <= left && right <= end) { maximum[node] += delta; lazy[node] += delta; return }
    const middle = (left + right) >>> 1
    if (start <= middle) addRange(node * 2, left, middle, start, end, delta)
    if (end > middle) addRange(node * 2 + 1, middle + 1, right, start, end, delta)
    maximum[node] = lazy[node] + Math.max(maximum[node * 2], maximum[node * 2 + 1])
  }
  const rangeMaximum = (node: number, left: number, right: number, start: number, end: number, inherited: number): number => {
    if (start <= left && right <= end) return inherited + maximum[node]
    const middle = (left + right) >>> 1, next = inherited + lazy[node]
    let value = 0
    if (start <= middle) value = rangeMaximum(node * 2, left, middle, start, end, next)
    if (end > middle) value = Math.max(value, rangeMaximum(node * 2 + 1, middle + 1, right, start, end, next))
    return value
  }
  for (const event of events) {
    if (event.delta > 0 && rangeMaximum(1, 0, EXCEL_MAX_COLUMNS - 1, event.start, event.end, 0) > 0) return true
    addRange(1, 0, EXCEL_MAX_COLUMNS - 1, event.start, event.end, event.delta)
  }
  return false
}

function canonicalOpcPartKey(value: string): string | undefined {
  if (!value || /[?#\\\u0000-\u001f\u007f]/.test(value) || /^[A-Za-z][A-Za-z0-9+.-]*:/.test(value) || /%(?:2f|5c)/i.test(value)) return undefined
  let decoded: string
  try { decoded = decodeURIComponent(value) } catch { return undefined }
  decoded = decoded.startsWith('/') ? decoded.slice(1) : decoded
  if (!decoded || decoded.startsWith('/') || decoded.endsWith('/') || decoded.includes('//')) return undefined
  const segments = decoded.split('/')
  if (segments.some((part) => !part || part === '.' || part === '..' || part.endsWith('.') || /[\\\u0000-\u001f\u007f]/.test(part))) return undefined
  return asciiLower(decoded)
}
function asciiLower(value: string): string { return value.replace(/[A-Z]/g, (item) => item.toLowerCase()) }
function nativeSheetNameCaseKey(value: string): string {
  let key = ''
  for (const character of value) {
    if (character === '\u0130') { key += 'i'; continue }
    const upper = [...character.toUpperCase()]
    const lower = upper.length === 1 ? [...upper[0].toLowerCase()] : [...character.toLowerCase()]
    key += lower.length === 1 ? lower[0] : character
  }
  return key
}
function isCanonicalSheetID(value: string): boolean { return /^(?:[1-9][0-9]{0,9})$/.test(value) && Number(value) <= 0xffffffff }
function finiteNumberLexical(value: string): boolean { return /^[+-]?(?:(?:[0-9]+(?:\.[0-9]*)?)|(?:\.[0-9]+))(?:[eE][+-]?[0-9]+)?$/.test(value) && Number.isFinite(Number(value)) }

function validIsoDateTime(value: string): boolean {
  const match = /^(\d{4})-(\d{2})-(\d{2})(?:(?:T(\d{2}):(\d{2}):(\d{2})(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})?)|(?:Z|[+-]\d{2}:\d{2}))?$/.exec(value)
  if (!match) return false
  const year = Number(match[1]), month = Number(match[2]), day = Number(match[3])
  const days = [31, leap(year) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
  const offset = /([+-])(\d{2}):(\d{2})$/.exec(value)
  return month >= 1 && month <= 12 && day >= 1 && day <= days[month - 1] &&
    (match[4] === undefined || (Number(match[4]) <= 23 && Number(match[5]) <= 59 && Number(match[6]) <= 59)) &&
    (!offset || (Number(offset[2]) <= 24 && Number(offset[3]) <= 59 && (Number(offset[2]) !== 24 || Number(offset[3]) === 0)))
}
function leap(year: number): boolean { return year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0) }

function decodeSpreadsheetString(value: string): string | undefined {
  let result = ''
  for (let index = 0; index < value.length;) {
    const unit = spreadsheetEscapeUnit(value, index)
    if (unit === undefined) { result += value[index++]; continue }
    index += 7
    if (unit >= 0xd800 && unit <= 0xdbff) {
      const lowUnit = spreadsheetEscapeUnit(value, index); if (lowUnit === undefined || lowUnit < 0xdc00 || lowUnit > 0xdfff) return undefined
      result += String.fromCodePoint(0x10000 + ((unit - 0xd800) << 10) + lowUnit - 0xdc00); index += 7
    } else if (unit >= 0xdc00 && unit <= 0xdfff) return undefined
    else result += String.fromCharCode(unit)
  }
  return result
}
function spreadsheetEscapeUnit(value: string, index: number): number | undefined {
  if (index + 7 > value.length || value[index] !== '_' || (value[index + 1] !== 'x' && value[index + 1] !== 'X') || value[index + 6] !== '_') return undefined
  const hexadecimal = value.slice(index + 2, index + 6)
  return /^[0-9A-Fa-f]{4}$/.test(hexadecimal) ? Number.parseInt(hexadecimal, 16) : undefined
}

function expectedSheetRefusal(reasons: Set<string>): string | undefined { return ['FORMULA_GROUPS', 'SHEET_PROTECTION', 'SHEET_DATA_ATTRIBUTES', 'UNSAFE_WORKSHEET_ATTRIBUTES', 'SHEET_DECLARATION_ATTRIBUTES'].find((item) => reasons.has(item)) }
function validUint32Lexical(value: string): boolean {
  if (!/^[0-9]+$/.test(value)) return false
  const significant = value.replace(/^0+/, '') || '0'
  return significant.length < 10 || (significant.length === 10 && significant <= '4294967295')
}
function nativeContractTextBytes(workbook: NativeWorkbookV1): number {
  let total = utf8ByteLength(workbook.document_id) + utf8ByteLength(workbook.source.workbook_part)
  const addText = (value: string | undefined) => { if (value !== undefined) total += utf8ByteLength(value) }
  const addValue = (value: NativeWorkbookValueV1 | undefined) => { if (!value) return; addText(value.kind); addText(value.storage); addText(value.lexical); addText(value.text) }
  for (const sheet of workbook.sheets) {
    addText(sheet.id); addText(sheet.name); addText(sheet.part_name); addText(sheet.refusal_code)
    for (const merged of sheet.merged_ranges) addText(merged.ref)
    for (const cell of sheet.cells) {
      addText(cell.ref); addValue(cell.value)
      if (cell.formula) { addText(cell.formula.text); addText(cell.formula.ref); addValue(cell.formula.cached) }
    }
  }
  for (const style of workbook.styles) {
    style.effective.unsupported.forEach(addText); addText(style.effective.number_format); addText(style.effective.font_name)
  }
  for (const capability of workbook.capabilities) { addText(capability.name); addText(capability.detail) }
  for (const part of workbook.passthrough_parts) { addText(part.part_name); addText(part.content_type) }
  for (const item of workbook.unsupported) {
    addText(item.id); addText(item.code); addText(item.capability); addText(item.scope_id); addText(item.preservation); addText(item.message)
    addText(item.part_name); addText(item.cell_ref); addText(item.range_ref)
  }
  return total
}
function add(issues: NativeWorkbookValidationIssue[], code: string, path: string, message: string): void { if (issues.length < XLSX_NATIVE_RESOURCE_LIMITS.maxIssues) issues.push({ code, path, message }) }

export type { NativeWorkbookValidationIssue, ValidateNativeWorkbookResult } from './nativeSchemaValidation.js'
