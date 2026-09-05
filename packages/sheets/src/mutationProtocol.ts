/**
 * Native XLSX save contract, version 1.
 *
 * This is deliberately not a serialized Univer mutation. UI command payloads
 * are renderer-specific and may contain workbook ids that change on reload.
 * This protocol uses only stable sheet ids and a compare-and-swap revision.
 */

export const WORKBOOK_MUTATION_PROTOCOL = 'injoffice.xlsx.mutations'
export const WORKBOOK_MUTATION_VERSION = 1 as const

export const EXCEL_MAX_ROWS = 1_048_576
export const EXCEL_MAX_COLUMNS = 16_384
export const MAX_WORKBOOK_MUTATIONS = 10_000

const MAX_REVISION_LENGTH = 256
const MAX_BATCH_ID_LENGTH = 128
const MAX_OPERATION_ID_LENGTH = 128
const MAX_SHEET_ID_LENGTH = 256
const MAX_CELL_TEXT_LENGTH = 32_767
const MAX_FORMULA_LENGTH = 8_192
const MAX_FORMAT_LENGTH = 255
const MAX_FONT_NAME_LENGTH = 255
const MAX_FONT_SIZE_POINTS = 409
const MAX_ROW_HEIGHT_POINTS = 409.5
const MAX_COLUMN_WIDTH = 255

export interface CellRef {
  /** Zero-based row index. */
  row: number
  /** Zero-based column index. */
  column: number
}

export interface RangeRef extends CellRef {
  /** Zero-based inclusive last row. */
  end_row: number
  /** Zero-based inclusive last column. */
  end_column: number
}

interface MutationBase {
  /** Stable within the request and across a safe retry. */
  operation_id: string
  /** Stable id imported from the workbook; never a tab name or index. */
  sheet_id: string
}

export interface CellSetValueMutation extends MutationBase {
  kind: 'cell.set_value'
  cell: CellRef
  /** Replaces formula content. Dates are serials plus a number-format style. */
  value: string | number | boolean
}

export interface CellClearValueMutation extends MutationBase {
  kind: 'cell.clear_value'
  cell: CellRef
}

export interface CellSetFormulaMutation extends MutationBase {
  kind: 'cell.set_formula'
  cell: CellRef
  /** Replaces literal content. A1 formula including the leading '='. */
  formula: string
}

export interface CellClearFormulaMutation extends MutationBase {
  kind: 'cell.clear_formula'
  cell: CellRef
}

export type HorizontalAlignment = 'general' | 'left' | 'center' | 'right'
export type VerticalAlignment = 'top' | 'middle' | 'bottom'

/**
 * Fixed v1 style vocabulary. Omitted means unchanged; null clears the direct
 * cell property so workbook defaults/inheritance apply. Arbitrary OOXML,
 * borders, rotations, indentation, and conditional formats are out of scope.
 */
export interface StyleDelta {
  number_format?: string | null
  font_name?: string | null
  font_size_points?: number | null
  bold?: boolean | null
  italic?: boolean | null
  font_color?: string | null
  fill_color?: string | null
  horizontal_alignment?: HorizontalAlignment | null
  vertical_alignment?: VerticalAlignment | null
  wrap_text?: boolean | null
}

export interface StylePatchMutation extends MutationBase {
  kind: 'style.patch'
  range: RangeRef
  style: StyleDelta
}

export interface RowSetHeightMutation extends MutationBase {
  kind: 'row.set_height'
  row: number
  /** Points. Zero hides the row. */
  height_points: number
}

export interface ColumnSetWidthMutation extends MutationBase {
  kind: 'column.set_width'
  column: number
  /** Excel character-width units. Zero hides the column. */
  width: number
}

export interface RangeMergeMutation extends MutationBase {
  kind: 'range.merge'
  range: RangeRef
}

export interface RangeUnmergeMutation extends MutationBase {
  kind: 'range.unmerge'
  range: RangeRef
}

export type SupportedWorkbookMutation =
  | CellSetValueMutation
  | CellClearValueMutation
  | CellSetFormulaMutation
  | CellClearFormulaMutation
  | StylePatchMutation
  | RowSetHeightMutation
  | ColumnSetWidthMutation
  | RangeMergeMutation
  | RangeUnmergeMutation

/** Recognized UI edits that v1 must refuse rather than omit from a save. */
export const UNSUPPORTED_STRUCTURAL_MUTATION_KINDS = [
  'sheet.add',
  'sheet.delete',
  'sheet.rename',
  'sheet.reorder',
  'row.insert',
  'row.delete',
  'row.move',
  'column.insert',
  'column.delete',
  'column.move',
  'range.move',
] as const

export type UnsupportedStructuralMutationKind = (typeof UNSUPPORTED_STRUCTURAL_MUTATION_KINDS)[number]

export interface WorkbookMutationBatchV1 {
  protocol: typeof WORKBOOK_MUTATION_PROTOCOL
  version: typeof WORKBOOK_MUTATION_VERSION
  /** Stable idempotency key for this logical save, retained across retries. */
  batch_id: string
  /** Opaque gateway-issued revision of the exact XLSX bytes being edited. */
  expected_revision: string
  /** Applied atomically and strictly in listed order. */
  operations: SupportedWorkbookMutation[]
}

export type WorkbookMutationIssueCode =
  | 'INVALID_ENVELOPE'
  | 'UNKNOWN_FIELD'
  | 'REQUIRED'
  | 'INVALID_TYPE'
  | 'OUT_OF_RANGE'
  | 'INVALID_VALUE'
  | 'UNSUPPORTED_PROTOCOL'
  | 'UNSUPPORTED_VERSION'
  | 'EMPTY_BATCH'
  | 'BATCH_TOO_LARGE'
  | 'DUPLICATE_OPERATION_ID'
  | 'UNKNOWN_OPERATION'
  | 'UNSUPPORTED_OPERATION'
  | 'EMPTY_STYLE_DELTA'
  | 'INVALID_RANGE'

/** JSON-serializable validation issue suitable for a 400 gateway response. */
export interface WorkbookMutationIssue {
  code: WorkbookMutationIssueCode
  /** RFC 6901 JSON Pointer into the submitted body. */
  path: string
  message: string
  operation_index?: number
  operation_id?: string
}

export type DecodeWorkbookMutationResult =
  | { ok: true; value: WorkbookMutationBatchV1 }
  | { ok: false; issues: WorkbookMutationIssue[] }

export class WorkbookMutationValidationError extends Error {
  readonly issues: WorkbookMutationIssue[]

  constructor(issues: WorkbookMutationIssue[]) {
    super(issues.map((issue) => `${issue.path || '/'}: ${issue.message}`).join('; '))
    this.name = 'WorkbookMutationValidationError'
    this.issues = issues
  }
}

type JsonObject = Record<string, unknown>

const supportedKinds = new Set<string>([
  'cell.set_value',
  'cell.clear_value',
  'cell.set_formula',
  'cell.clear_formula',
  'style.patch',
  'row.set_height',
  'column.set_width',
  'range.merge',
  'range.unmerge',
])
const unsupportedKinds = new Set<string>(UNSUPPORTED_STRUCTURAL_MUTATION_KINDS)

function isObject(value: unknown): value is JsonObject {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

function issue(
  issues: WorkbookMutationIssue[],
  code: WorkbookMutationIssueCode,
  path: string,
  message: string,
  operationIndex?: number,
  operationId?: string,
): void {
  const value: WorkbookMutationIssue = { code, path, message }
  if (operationIndex !== undefined) value.operation_index = operationIndex
  if (operationId !== undefined) value.operation_id = operationId
  issues.push(value)
}

function rejectUnknownFields(
  object: JsonObject,
  allowed: readonly string[],
  path: string,
  issues: WorkbookMutationIssue[],
  operationIndex?: number,
  operationId?: string,
): void {
  const allowedSet = new Set(allowed)
  for (const key of Object.keys(object).sort()) {
    if (!allowedSet.has(key)) issue(issues, 'UNKNOWN_FIELD', `${path}/${key}`, `unknown field ${JSON.stringify(key)}`, operationIndex, operationId)
  }
}

function requiredString(
  value: unknown,
  path: string,
  maxLength: number,
  issues: WorkbookMutationIssue[],
  operationIndex?: number,
  operationId?: string,
): string | null {
  if (value === undefined) {
    issue(issues, 'REQUIRED', path, 'field is required', operationIndex, operationId)
    return null
  }
  if (typeof value !== 'string') {
    issue(issues, 'INVALID_TYPE', path, 'must be a string', operationIndex, operationId)
    return null
  }
  if (value.length === 0 || value.trim() !== value) {
    issue(issues, 'INVALID_VALUE', path, 'must be non-empty with no leading or trailing whitespace', operationIndex, operationId)
    return null
  }
  if (value.length > maxLength) {
    issue(issues, 'OUT_OF_RANGE', path, `must be at most ${maxLength} UTF-16 code units`, operationIndex, operationId)
    return null
  }
  return value
}

function restrictedIdentifier(
  value: unknown,
  path: string,
  maxLength: number,
  issues: WorkbookMutationIssue[],
  operationIndex?: number,
): string | null {
  const parsed = requiredString(value, path, maxLength, issues, operationIndex)
  if (parsed !== null && !/^[A-Za-z0-9][A-Za-z0-9._:-]*$/.test(parsed)) {
    issue(issues, 'INVALID_VALUE', path, 'must use only ASCII letters, digits, dot, underscore, colon, or hyphen', operationIndex)
    return null
  }
  return parsed
}

function boundedNumber(
  value: unknown,
  path: string,
  min: number,
  max: number,
  integer: boolean,
  issues: WorkbookMutationIssue[],
  operationIndex?: number,
  operationId?: string,
): number | null {
  if (value === undefined) {
    issue(issues, 'REQUIRED', path, 'field is required', operationIndex, operationId)
    return null
  }
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    issue(issues, 'INVALID_TYPE', path, 'must be a finite number', operationIndex, operationId)
    return null
  }
  if ((integer && !Number.isInteger(value)) || value < min || value > max) {
    issue(issues, 'OUT_OF_RANGE', path, `must be ${integer ? 'an integer ' : ''}between ${min} and ${max}`, operationIndex, operationId)
    return null
  }
  return value
}

function parseCellRef(value: unknown, path: string, issues: WorkbookMutationIssue[], operationIndex: number, operationId?: string): CellRef | null {
  if (!isObject(value)) {
    issue(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, value === undefined ? 'field is required' : 'must be an object', operationIndex, operationId)
    return null
  }
  rejectUnknownFields(value, ['row', 'column'], path, issues, operationIndex, operationId)
  const row = boundedNumber(value.row, `${path}/row`, 0, EXCEL_MAX_ROWS - 1, true, issues, operationIndex, operationId)
  const column = boundedNumber(value.column, `${path}/column`, 0, EXCEL_MAX_COLUMNS - 1, true, issues, operationIndex, operationId)
  return row === null || column === null ? null : { row, column }
}

function parseRangeRef(value: unknown, path: string, issues: WorkbookMutationIssue[], operationIndex: number, operationId?: string): RangeRef | null {
  if (!isObject(value)) {
    issue(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, value === undefined ? 'field is required' : 'must be an object', operationIndex, operationId)
    return null
  }
  rejectUnknownFields(value, ['row', 'column', 'end_row', 'end_column'], path, issues, operationIndex, operationId)
  const row = boundedNumber(value.row, `${path}/row`, 0, EXCEL_MAX_ROWS - 1, true, issues, operationIndex, operationId)
  const column = boundedNumber(value.column, `${path}/column`, 0, EXCEL_MAX_COLUMNS - 1, true, issues, operationIndex, operationId)
  const endRow = boundedNumber(value.end_row, `${path}/end_row`, 0, EXCEL_MAX_ROWS - 1, true, issues, operationIndex, operationId)
  const endColumn = boundedNumber(value.end_column, `${path}/end_column`, 0, EXCEL_MAX_COLUMNS - 1, true, issues, operationIndex, operationId)
  if (row === null || column === null || endRow === null || endColumn === null) return null
  if (endRow < row || endColumn < column) {
    issue(issues, 'INVALID_RANGE', path, 'range end must not precede its start', operationIndex, operationId)
    return null
  }
  return { row, column, end_row: endRow, end_column: endColumn }
}

const styleKeys = [
  'number_format',
  'font_name',
  'font_size_points',
  'bold',
  'italic',
  'font_color',
  'fill_color',
  'horizontal_alignment',
  'vertical_alignment',
  'wrap_text',
] as const

function nullableString(
  value: unknown,
  path: string,
  maxLength: number,
  issues: WorkbookMutationIssue[],
  operationIndex: number,
  operationId?: string,
): string | null | undefined {
  if (value === undefined || value === null) return value
  if (typeof value !== 'string') {
    issue(issues, 'INVALID_TYPE', path, 'must be a string or null', operationIndex, operationId)
    return undefined
  }
  if (value.length === 0 || value.length > maxLength) {
    issue(issues, 'OUT_OF_RANGE', path, `must contain 1 to ${maxLength} UTF-16 code units`, operationIndex, operationId)
    return undefined
  }
  return value
}

function nullableBoolean(
  value: unknown,
  path: string,
  issues: WorkbookMutationIssue[],
  operationIndex: number,
  operationId?: string,
): boolean | null | undefined {
  if (value === undefined || value === null) return value
  if (typeof value !== 'boolean') {
    issue(issues, 'INVALID_TYPE', path, 'must be a boolean or null', operationIndex, operationId)
    return undefined
  }
  return value
}

function nullableEnum<T extends string>(
  value: unknown,
  allowed: readonly T[],
  path: string,
  issues: WorkbookMutationIssue[],
  operationIndex: number,
  operationId?: string,
): T | null | undefined {
  if (value === undefined || value === null) return value
  if (typeof value !== 'string' || !allowed.includes(value as T)) {
    issue(issues, 'INVALID_VALUE', path, `must be one of ${allowed.join(', ')} or null`, operationIndex, operationId)
    return undefined
  }
  return value as T
}

function nullableColor(
  value: unknown,
  path: string,
  issues: WorkbookMutationIssue[],
  operationIndex: number,
  operationId?: string,
): string | null | undefined {
  if (value === undefined || value === null) return value
  if (typeof value !== 'string' || !/^#[0-9A-F]{6}$/.test(value)) {
    issue(issues, 'INVALID_VALUE', path, 'must be canonical uppercase #RRGGBB or null', operationIndex, operationId)
    return undefined
  }
  return value
}

function parseStyleDelta(value: unknown, path: string, issues: WorkbookMutationIssue[], operationIndex: number, operationId?: string): StyleDelta | null {
  if (!isObject(value)) {
    issue(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, value === undefined ? 'field is required' : 'must be an object', operationIndex, operationId)
    return null
  }
  rejectUnknownFields(value, styleKeys, path, issues, operationIndex, operationId)
  const present = styleKeys.filter((key) => Object.prototype.hasOwnProperty.call(value, key))
  if (present.length === 0) {
    issue(issues, 'EMPTY_STYLE_DELTA', path, 'must contain at least one supported style field', operationIndex, operationId)
    return null
  }

  const out: StyleDelta = {}
  const put = <K extends keyof StyleDelta>(key: K, parsed: StyleDelta[K] | undefined): void => {
    if (Object.prototype.hasOwnProperty.call(value, key) && parsed !== undefined) out[key] = parsed
  }
  put('number_format', nullableString(value.number_format, `${path}/number_format`, MAX_FORMAT_LENGTH, issues, operationIndex, operationId))
  put('font_name', nullableString(value.font_name, `${path}/font_name`, MAX_FONT_NAME_LENGTH, issues, operationIndex, operationId))
  let fontSize: number | null | undefined = value.font_size_points as number | null | undefined
  if (fontSize !== undefined && fontSize !== null) {
    fontSize = boundedNumber(fontSize, `${path}/font_size_points`, 1, MAX_FONT_SIZE_POINTS, false, issues, operationIndex, operationId) ?? undefined
  }
  put('font_size_points', fontSize)
  put('bold', nullableBoolean(value.bold, `${path}/bold`, issues, operationIndex, operationId))
  put('italic', nullableBoolean(value.italic, `${path}/italic`, issues, operationIndex, operationId))
  put('font_color', nullableColor(value.font_color, `${path}/font_color`, issues, operationIndex, operationId))
  put('fill_color', nullableColor(value.fill_color, `${path}/fill_color`, issues, operationIndex, operationId))
  put('horizontal_alignment', nullableEnum(value.horizontal_alignment, ['general', 'left', 'center', 'right'], `${path}/horizontal_alignment`, issues, operationIndex, operationId))
  put('vertical_alignment', nullableEnum(value.vertical_alignment, ['top', 'middle', 'bottom'], `${path}/vertical_alignment`, issues, operationIndex, operationId))
  put('wrap_text', nullableBoolean(value.wrap_text, `${path}/wrap_text`, issues, operationIndex, operationId))
  return out
}

function parseOperation(value: unknown, index: number, issues: WorkbookMutationIssue[]): SupportedWorkbookMutation | null {
  const path = `/operations/${index}`
  if (!isObject(value)) {
    issue(issues, 'INVALID_TYPE', path, 'operation must be an object', index)
    return null
  }

  const operationId = restrictedIdentifier(value.operation_id, `${path}/operation_id`, MAX_OPERATION_ID_LENGTH, issues, index) ?? undefined
  const kind = typeof value.kind === 'string' ? value.kind : undefined
  if (kind === undefined) {
    issue(issues, value.kind === undefined ? 'REQUIRED' : 'INVALID_TYPE', `${path}/kind`, value.kind === undefined ? 'field is required' : 'must be a string', index, operationId)
    return null
  }
  if (unsupportedKinds.has(kind)) {
    issue(issues, 'UNSUPPORTED_OPERATION', `${path}/kind`, `structural operation ${JSON.stringify(kind)} is not supported by protocol v1`, index, operationId)
    return null
  }
  if (!supportedKinds.has(kind)) {
    issue(issues, 'UNKNOWN_OPERATION', `${path}/kind`, `unknown operation ${JSON.stringify(kind)}`, index, operationId)
    return null
  }

  const sheetId = requiredString(value.sheet_id, `${path}/sheet_id`, MAX_SHEET_ID_LENGTH, issues, index, operationId)
  const base = operationId && sheetId ? { operation_id: operationId, sheet_id: sheetId } : null
  switch (kind) {
    case 'cell.set_value': {
      rejectUnknownFields(value, ['operation_id', 'kind', 'sheet_id', 'cell', 'value'], path, issues, index, operationId)
      const cell = parseCellRef(value.cell, `${path}/cell`, issues, index, operationId)
      const cellValue = value.value
      if (cellValue === undefined) issue(issues, 'REQUIRED', `${path}/value`, 'field is required', index, operationId)
      else if (typeof cellValue === 'number' && !Number.isFinite(cellValue)) issue(issues, 'INVALID_TYPE', `${path}/value`, 'number value must be finite', index, operationId)
      else if (!['string', 'number', 'boolean'].includes(typeof cellValue)) issue(issues, 'INVALID_TYPE', `${path}/value`, 'must be a string, finite number, or boolean', index, operationId)
      else if (typeof cellValue === 'string' && cellValue.length > MAX_CELL_TEXT_LENGTH) issue(issues, 'OUT_OF_RANGE', `${path}/value`, `string value must be at most ${MAX_CELL_TEXT_LENGTH} UTF-16 code units`, index, operationId)
      return base && cell && (typeof cellValue === 'string' || typeof cellValue === 'boolean' || (typeof cellValue === 'number' && Number.isFinite(cellValue)))
        ? { ...base, kind, cell, value: cellValue }
        : null
    }
    case 'cell.clear_value':
    case 'cell.clear_formula': {
      rejectUnknownFields(value, ['operation_id', 'kind', 'sheet_id', 'cell'], path, issues, index, operationId)
      const cell = parseCellRef(value.cell, `${path}/cell`, issues, index, operationId)
      return base && cell ? { ...base, kind, cell } : null
    }
    case 'cell.set_formula': {
      rejectUnknownFields(value, ['operation_id', 'kind', 'sheet_id', 'cell', 'formula'], path, issues, index, operationId)
      const cell = parseCellRef(value.cell, `${path}/cell`, issues, index, operationId)
      const formula = requiredString(value.formula, `${path}/formula`, MAX_FORMULA_LENGTH, issues, index, operationId)
      if (formula !== null && !formula.startsWith('=')) issue(issues, 'INVALID_VALUE', `${path}/formula`, "must start with '='", index, operationId)
      return base && cell && formula?.startsWith('=') ? { ...base, kind, cell, formula } : null
    }
    case 'style.patch': {
      rejectUnknownFields(value, ['operation_id', 'kind', 'sheet_id', 'range', 'style'], path, issues, index, operationId)
      const range = parseRangeRef(value.range, `${path}/range`, issues, index, operationId)
      const style = parseStyleDelta(value.style, `${path}/style`, issues, index, operationId)
      return base && range && style ? { ...base, kind, range, style } : null
    }
    case 'row.set_height': {
      rejectUnknownFields(value, ['operation_id', 'kind', 'sheet_id', 'row', 'height_points'], path, issues, index, operationId)
      const row = boundedNumber(value.row, `${path}/row`, 0, EXCEL_MAX_ROWS - 1, true, issues, index, operationId)
      const heightPoints = boundedNumber(value.height_points, `${path}/height_points`, 0, MAX_ROW_HEIGHT_POINTS, false, issues, index, operationId)
      return base && row !== null && heightPoints !== null ? { ...base, kind, row, height_points: heightPoints } : null
    }
    case 'column.set_width': {
      rejectUnknownFields(value, ['operation_id', 'kind', 'sheet_id', 'column', 'width'], path, issues, index, operationId)
      const column = boundedNumber(value.column, `${path}/column`, 0, EXCEL_MAX_COLUMNS - 1, true, issues, index, operationId)
      const width = boundedNumber(value.width, `${path}/width`, 0, MAX_COLUMN_WIDTH, false, issues, index, operationId)
      return base && column !== null && width !== null ? { ...base, kind, column, width } : null
    }
    case 'range.merge':
    case 'range.unmerge': {
      rejectUnknownFields(value, ['operation_id', 'kind', 'sheet_id', 'range'], path, issues, index, operationId)
      const range = parseRangeRef(value.range, `${path}/range`, issues, index, operationId)
      if (range && range.row === range.end_row && range.column === range.end_column) {
        issue(issues, 'INVALID_RANGE', `${path}/range`, `${kind} requires at least two cells`, index, operationId)
        return null
      }
      return base && range ? { ...base, kind, range } : null
    }
  }
  return null
}

/** Validate untrusted JSON and rebuild it in canonical field order. */
export function decodeWorkbookMutationBatch(value: unknown): DecodeWorkbookMutationResult {
  const issues: WorkbookMutationIssue[] = []
  if (!isObject(value)) return { ok: false, issues: [{ code: 'INVALID_ENVELOPE', path: '', message: 'request body must be an object' }] }

  rejectUnknownFields(value, ['protocol', 'version', 'batch_id', 'expected_revision', 'operations'], '', issues)
  if (value.protocol === undefined) {
    issue(issues, 'REQUIRED', '/protocol', 'field is required')
  } else if (value.protocol !== WORKBOOK_MUTATION_PROTOCOL) {
    issue(issues, 'UNSUPPORTED_PROTOCOL', '/protocol', `must equal ${JSON.stringify(WORKBOOK_MUTATION_PROTOCOL)}`)
  }
  if (value.version === undefined) {
    issue(issues, 'REQUIRED', '/version', 'field is required')
  } else if (value.version !== WORKBOOK_MUTATION_VERSION) {
    issue(issues, 'UNSUPPORTED_VERSION', '/version', `must equal ${WORKBOOK_MUTATION_VERSION}`)
  }
  const batchId = restrictedIdentifier(value.batch_id, '/batch_id', MAX_BATCH_ID_LENGTH, issues)
  const expectedRevision = requiredString(value.expected_revision, '/expected_revision', MAX_REVISION_LENGTH, issues)

  const operations: SupportedWorkbookMutation[] = []
  if (!Array.isArray(value.operations)) {
    issue(issues, value.operations === undefined ? 'REQUIRED' : 'INVALID_TYPE', '/operations', value.operations === undefined ? 'field is required' : 'must be an array')
  } else {
    if (value.operations.length === 0) {
      issue(issues, 'EMPTY_BATCH', '/operations', 'must contain at least one operation')
    } else if (value.operations.length > MAX_WORKBOOK_MUTATIONS) {
      // Refuse before visiting any entries. The body is untrusted, and both
      // validation time and the number of returned issues must stay bounded.
      issue(issues, 'BATCH_TOO_LARGE', '/operations', `must contain at most ${MAX_WORKBOOK_MUTATIONS} operations`)
    } else {
      const ids = new Map<string, number>()
      for (let index = 0; index < value.operations.length; index++) {
        const operation = parseOperation(value.operations[index], index, issues)
        if (!operation) continue
        const previous = ids.get(operation.operation_id)
        if (previous !== undefined) {
          issue(issues, 'DUPLICATE_OPERATION_ID', `/operations/${index}/operation_id`, `duplicates operation id at index ${previous}`, index, operation.operation_id)
        } else {
          ids.set(operation.operation_id, index)
        }
        operations.push(operation)
      }
    }
  }

  if (issues.length > 0 || batchId === null || expectedRevision === null) return { ok: false, issues }
  return {
    ok: true,
    value: {
      protocol: WORKBOOK_MUTATION_PROTOCOL,
      version: WORKBOOK_MUTATION_VERSION,
      batch_id: batchId,
      expected_revision: expectedRevision,
      operations,
    },
  }
}

/** Validate and serialize with stable field ordering. Operation order is never sorted. */
export function encodeWorkbookMutationBatch(batch: WorkbookMutationBatchV1): string {
  const decoded = decodeWorkbookMutationBatch(batch)
  if (!decoded.ok) throw new WorkbookMutationValidationError(decoded.issues)
  return JSON.stringify(decoded.value)
}
