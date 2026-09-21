/**
 * Chart insert/update/delete operations for workbook mutation protocol v1.
 *
 * Kept in this file so concurrent sheets lanes can rebase mutationProtocol.ts
 * without owning the chart vocabulary. Excel-openable bar, column, line, and
 * pie charts over a selected literal range.
 */

import type { RangeRef, WorkbookMutationIssue, WorkbookMutationIssueCode } from './mutationProtocol.js'

const EXCEL_MAX_ROWS = 1_048_576
const EXCEL_MAX_COLUMNS = 16_384
const MAX_TITLE_LENGTH = 1_024
const MAX_CHART_CATEGORIES = 1_000
const MAX_CHART_SERIES = 8

export const CHART_MUTATION_KINDS = ['chart.insert', 'chart.update', 'chart.delete'] as const
export type ChartMutationKind = (typeof CHART_MUTATION_KINDS)[number]
export const CHART_TYPES = ['column', 'bar', 'line', 'pie'] as const
export type ChartTypeV1 = (typeof CHART_TYPES)[number]

/** DrawingML two-cell anchor. Coordinates are zero-based; to is exclusive. */
export interface ChartAnchor {
  from_row: number
  from_column: number
  to_row: number
  to_column: number
}

/** Stable native chart identity. Field names match Go xlsxpatch.ChartIdentity. */
export interface ChartIdentity {
  part: string
  drawingPart: string
  objectId: number
}

interface MutationBase {
  operation_id: string
  sheet_id: string
}

export interface ChartInsertMutation extends MutationBase {
  kind: 'chart.insert'
  chart_type: ChartTypeV1
  title: string
  range: RangeRef
  anchor: ChartAnchor
}

export interface ChartUpdateMutation extends MutationBase {
  kind: 'chart.update'
  identity: ChartIdentity
  expected_fingerprint_sha256: string
  chart_type: ChartTypeV1
  title: string
  range: RangeRef
  anchor: ChartAnchor
}

export interface ChartDeleteMutation extends MutationBase {
  kind: 'chart.delete'
  identity: ChartIdentity
  expected_fingerprint_sha256: string
}

export type ChartMutation = ChartInsertMutation | ChartUpdateMutation | ChartDeleteMutation

type JsonObject = Record<string, unknown>

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

function boundedInteger(
  value: unknown,
  path: string,
  min: number,
  max: number,
  issues: WorkbookMutationIssue[],
  operationIndex?: number,
  operationId?: string,
): number | null {
  if (value === undefined) {
    issue(issues, 'REQUIRED', path, 'field is required', operationIndex, operationId)
    return null
  }
  if (typeof value !== 'number' || !Number.isFinite(value) || !Number.isInteger(value) || value < min || value > max) {
    issue(issues, typeof value === 'number' ? 'OUT_OF_RANGE' : 'INVALID_TYPE', path, `must be an integer between ${min} and ${max}`, operationIndex, operationId)
    return null
  }
  return value
}

function parseRange(value: unknown, path: string, issues: WorkbookMutationIssue[], operationIndex: number, operationId?: string): RangeRef | null {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    issue(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, value === undefined ? 'field is required' : 'must be an object', operationIndex, operationId)
    return null
  }
  const object = value as JsonObject
  rejectUnknownFields(object, ['row', 'column', 'end_row', 'end_column'], path, issues, operationIndex, operationId)
  const row = boundedInteger(object.row, `${path}/row`, 0, EXCEL_MAX_ROWS - 1, issues, operationIndex, operationId)
  const column = boundedInteger(object.column, `${path}/column`, 0, EXCEL_MAX_COLUMNS - 1, issues, operationIndex, operationId)
  const endRow = boundedInteger(object.end_row, `${path}/end_row`, 0, EXCEL_MAX_ROWS - 1, issues, operationIndex, operationId)
  const endColumn = boundedInteger(object.end_column, `${path}/end_column`, 0, EXCEL_MAX_COLUMNS - 1, issues, operationIndex, operationId)
  if (row === null || column === null || endRow === null || endColumn === null) return null
  if (endRow < row || endColumn < column) {
    issue(issues, 'INVALID_RANGE', path, 'range end must not precede its start', operationIndex, operationId)
    return null
  }
  if (endRow === row || endColumn === column) {
    issue(issues, 'INVALID_RANGE', path, 'chart source requires a header row, category column, and at least one numeric series', operationIndex, operationId)
    return null
  }
  if (endRow - row > MAX_CHART_CATEGORIES || endColumn - column > MAX_CHART_SERIES) {
    issue(issues, 'OUT_OF_RANGE', path, `charts support up to ${MAX_CHART_CATEGORIES} categories and ${MAX_CHART_SERIES} numeric series`, operationIndex, operationId)
    return null
  }
  return { row, column, end_row: endRow, end_column: endColumn }
}

function parseAnchor(value: unknown, path: string, issues: WorkbookMutationIssue[], operationIndex: number, operationId?: string): ChartAnchor | null {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    issue(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, value === undefined ? 'field is required' : 'must be an object', operationIndex, operationId)
    return null
  }
  const object = value as JsonObject
  rejectUnknownFields(object, ['from_row', 'from_column', 'to_row', 'to_column'], path, issues, operationIndex, operationId)
  const fromRow = boundedInteger(object.from_row, `${path}/from_row`, 0, EXCEL_MAX_ROWS - 1, issues, operationIndex, operationId)
  const fromColumn = boundedInteger(object.from_column, `${path}/from_column`, 0, EXCEL_MAX_COLUMNS - 1, issues, operationIndex, operationId)
  const toRow = boundedInteger(object.to_row, `${path}/to_row`, 0, EXCEL_MAX_ROWS, issues, operationIndex, operationId)
  const toColumn = boundedInteger(object.to_column, `${path}/to_column`, 0, EXCEL_MAX_COLUMNS, issues, operationIndex, operationId)
  if (fromRow === null || fromColumn === null || toRow === null || toColumn === null) return null
  if (toRow <= fromRow || toColumn <= fromColumn) {
    issue(issues, 'INVALID_RANGE', path, 'chart anchor to cell must be below and right of from', operationIndex, operationId)
    return null
  }
  return { from_row: fromRow, from_column: fromColumn, to_row: toRow, to_column: toColumn }
}

function parseIdentity(value: unknown, path: string, issues: WorkbookMutationIssue[], operationIndex: number, operationId?: string): ChartIdentity | null {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    issue(issues, value === undefined ? 'REQUIRED' : 'INVALID_TYPE', path, value === undefined ? 'field is required' : 'must be an object', operationIndex, operationId)
    return null
  }
  const object = value as JsonObject
  rejectUnknownFields(object, ['part', 'drawingPart', 'objectId'], path, issues, operationIndex, operationId)
  const part = typeof object.part === 'string' ? object.part : null
  const drawingPart = typeof object.drawingPart === 'string' ? object.drawingPart : null
  if (object.part === undefined) issue(issues, 'REQUIRED', `${path}/part`, 'field is required', operationIndex, operationId)
  else if (part === null) issue(issues, 'INVALID_TYPE', `${path}/part`, 'must be a string', operationIndex, operationId)
  else if (!/^xl\/charts\/[^/]+\.xml$/.test(part)) issue(issues, 'INVALID_VALUE', `${path}/part`, 'must be an xl/charts/*.xml part', operationIndex, operationId)
  if (object.drawingPart === undefined) issue(issues, 'REQUIRED', `${path}/drawingPart`, 'field is required', operationIndex, operationId)
  else if (drawingPart === null) issue(issues, 'INVALID_TYPE', `${path}/drawingPart`, 'must be a string', operationIndex, operationId)
  else if (!/^xl\/drawings\/[^/]+\.xml$/.test(drawingPart)) issue(issues, 'INVALID_VALUE', `${path}/drawingPart`, 'must be an xl/drawings/*.xml part', operationIndex, operationId)
  const objectId = boundedInteger(object.objectId, `${path}/objectId`, 1, 0xffffffff, issues, operationIndex, operationId)
  if (part === null || drawingPart === null || objectId === null || !/^xl\/charts\/[^/]+\.xml$/.test(part) || !/^xl\/drawings\/[^/]+\.xml$/.test(drawingPart)) return null
  return { part, drawingPart, objectId }
}

function parseFingerprint(value: unknown, path: string, issues: WorkbookMutationIssue[], operationIndex: number, operationId?: string): string | null {
  if (value === undefined) {
    issue(issues, 'REQUIRED', path, 'field is required', operationIndex, operationId)
    return null
  }
  if (typeof value !== 'string' || !/^sha256:[a-f0-9]{64}$/.test(value)) {
    issue(issues, typeof value === 'string' ? 'INVALID_VALUE' : 'INVALID_TYPE', path, 'must be sha256 followed by a lowercase 64-hex digest', operationIndex, operationId)
    return null
  }
  return value
}

function parseChartType(value: unknown, path: string, issues: WorkbookMutationIssue[], operationIndex: number, operationId?: string): ChartTypeV1 | null {
  if (value === undefined) {
    issue(issues, 'REQUIRED', path, 'field is required', operationIndex, operationId)
    return null
  }
  if (typeof value !== 'string' || !CHART_TYPES.includes(value as ChartTypeV1)) {
    issue(issues, 'INVALID_VALUE', path, `must be one of ${CHART_TYPES.join(', ')}`, operationIndex, operationId)
    return null
  }
  return value as ChartTypeV1
}

function parseTitle(value: unknown, path: string, issues: WorkbookMutationIssue[], operationIndex: number, operationId?: string): string | null {
  if (value === undefined) {
    issue(issues, 'REQUIRED', path, 'field is required', operationIndex, operationId)
    return null
  }
  if (typeof value !== 'string') {
    issue(issues, 'INVALID_TYPE', path, 'must be a string', operationIndex, operationId)
    return null
  }
  if (value.length > MAX_TITLE_LENGTH) {
    issue(issues, 'OUT_OF_RANGE', path, `must be at most ${MAX_TITLE_LENGTH} UTF-16 code units`, operationIndex, operationId)
    return null
  }
  return value
}

/** Parse one chart.* operation. Caller has already accepted the kind. */
export function parseChartMutation(
  value: JsonObject,
  path: string,
  index: number,
  operationId: string | undefined,
  sheetId: string | null,
  issues: WorkbookMutationIssue[],
): ChartMutation | null {
  const kind = value.kind as ChartMutationKind
  if (kind === 'chart.insert') {
    rejectUnknownFields(value, ['operation_id', 'kind', 'sheet_id', 'chart_type', 'title', 'range', 'anchor'], path, issues, index, operationId)
    const chartType = parseChartType(value.chart_type, `${path}/chart_type`, issues, index, operationId)
    const title = parseTitle(value.title, `${path}/title`, issues, index, operationId)
    const range = parseRange(value.range, `${path}/range`, issues, index, operationId)
    const anchor = parseAnchor(value.anchor, `${path}/anchor`, issues, index, operationId)
    return operationId && sheetId && chartType && title !== null && range && anchor
      ? { operation_id: operationId, sheet_id: sheetId, kind, chart_type: chartType, title, range, anchor }
      : null
  }
  if (kind === 'chart.update') {
    rejectUnknownFields(value, ['operation_id', 'kind', 'sheet_id', 'identity', 'expected_fingerprint_sha256', 'chart_type', 'title', 'range', 'anchor'], path, issues, index, operationId)
    const identity = parseIdentity(value.identity, `${path}/identity`, issues, index, operationId)
    const fingerprint = parseFingerprint(value.expected_fingerprint_sha256, `${path}/expected_fingerprint_sha256`, issues, index, operationId)
    const chartType = parseChartType(value.chart_type, `${path}/chart_type`, issues, index, operationId)
    const title = parseTitle(value.title, `${path}/title`, issues, index, operationId)
    const range = parseRange(value.range, `${path}/range`, issues, index, operationId)
    const anchor = parseAnchor(value.anchor, `${path}/anchor`, issues, index, operationId)
    return operationId && sheetId && identity && fingerprint && chartType && title !== null && range && anchor
      ? { operation_id: operationId, sheet_id: sheetId, kind, identity, expected_fingerprint_sha256: fingerprint, chart_type: chartType, title, range, anchor }
      : null
  }
  rejectUnknownFields(value, ['operation_id', 'kind', 'sheet_id', 'identity', 'expected_fingerprint_sha256'], path, issues, index, operationId)
  const identity = parseIdentity(value.identity, `${path}/identity`, issues, index, operationId)
  const fingerprint = parseFingerprint(value.expected_fingerprint_sha256, `${path}/expected_fingerprint_sha256`, issues, index, operationId)
  return operationId && sheetId && identity && fingerprint
    ? { operation_id: operationId, sheet_id: sheetId, kind: 'chart.delete', identity, expected_fingerprint_sha256: fingerprint }
    : null
}
