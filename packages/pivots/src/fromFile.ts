import { AGG_KINDS, type AggKind, type CellRangeRef, type NativePivotIdentity, type PivotFieldMember, type PivotSpec } from './types'

/** JSON shape returned by go/xlsxpatch.ReadPivots. */
export interface FilePivotInfo {
  part: string
  identity?: NativePivotIdentity
  name?: string
  cacheId: number
  sourceSheetName?: string
  sourceRef?: string
  fields: string[]
  fieldMembers?: Array<{ field: string; items: PivotFieldMember[] }>
  targetSheetName?: string
  targetRef?: string
  rowFields: string[]
  colFields: string[]
  dataFields: Array<{ field: string; agg: string }>
  pageFields?: Array<{ field: string; selectedItem?: string }>
  memberFilters?: Array<{ field: string; excludedItems: string[] }>
  sorts?: Array<{ field: string; direction: 'ascending' | 'descending' }>
  grandTotals: boolean
  warnings?: string[]
}

export interface PivotFileContext {
  sheetIdOf: (sheetName: string) => string | null
}

export interface PivotFileResult {
  pivots: PivotSpec[]
  skipped: string[]
}

function columnIndex(label: string): number {
  let value = 0
  for (const char of label.toUpperCase()) value = value * 26 + char.charCodeAt(0) - 64
  return value - 1
}

function parseA1Range(value: string, sheetId: string): CellRangeRef | null {
  const match = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i.exec(value.trim())
  if (!match) return null
  const startRow = Number(match[2]) - 1
  const startColumn = columnIndex(match[1])
  const endRow = Number(match[4] ?? match[2]) - 1
  const endColumn = columnIndex(match[3] ?? match[1])
  if (startRow < 0 || startColumn < 0 || endRow < startRow || endColumn < startColumn) return null
  return { sheetId, startRow, startColumn, endRow, endColumn }
}

function labelOf(info: FilePivotInfo): string {
  return info.name?.trim() || info.part
}

/**
 * Convert hydrated native pivot metadata to editable PivotSpecs. Conversion
 * is deliberately fail-closed: a native feature the model cannot represent
 * is reported in skipped instead of being silently discarded on the next save.
 */
export function pivotsFromFile(infos: FilePivotInfo[], context: PivotFileContext): PivotFileResult {
  const pivots: PivotSpec[] = []
  const skipped: string[] = []
  const ids = new Set<string>()
  for (const info of infos) {
    const label = labelOf(info)
    if (info.warnings?.length) {
      skipped.push(`pivot ${label}: ${info.warnings.join('; ')}`)
      continue
    }
    if (!info.sourceSheetName || !info.targetSheetName || !info.sourceRef || !info.targetRef) {
      skipped.push(`pivot ${label}: source or target range is missing`)
      continue
    }
    const sourceSheetId = context.sheetIdOf(info.sourceSheetName)
    const targetSheetId = context.sheetIdOf(info.targetSheetName)
    if (!sourceSheetId || !targetSheetId) {
      skipped.push(`pivot ${label}: source or target sheet is unavailable`)
      continue
    }
    const source = parseA1Range(info.sourceRef, sourceSheetId)
    const targetRange = parseA1Range(info.targetRef, targetSheetId)
    if (!source || !targetRange) {
      skipped.push(`pivot ${label}: source or target A1 reference is unsupported`)
      continue
    }
    const values: PivotSpec['values'] = []
    let unsupported = ''
    for (const field of info.dataFields) {
      if (!(AGG_KINDS as readonly string[]).includes(field.agg)) {
        unsupported = field.agg
        break
      }
      values.push({ field: field.field, agg: field.agg as AggKind })
    }
    if (unsupported) {
      skipped.push(`pivot ${label}: aggregation ${unsupported} is unsupported`)
      continue
    }
    if (values.length === 0) {
      skipped.push(`pivot ${label}: no value fields were hydrated`)
      continue
    }
    const baseId = (info.name || info.part.split('/').pop() || 'pivot').replace(/[^A-Za-z0-9_-]+/g, '_')
    let id = baseId
    for (let suffix = 2; ids.has(id); suffix++) id = `${baseId}_${suffix}`
    ids.add(id)
    const pivot: PivotSpec = {
      id,
      nativeIdentity: info.identity ?? { part: info.part },
      source,
      rows: [...info.rowFields],
      columns: [...info.colFields],
      values,
      target: { sheetId: targetSheetId, startRow: targetRange.startRow, startColumn: targetRange.startColumn },
      grandTotals: info.grandTotals,
    }
    if (info.pageFields?.length) pivot.pageFields = info.pageFields.map((page) => ({ ...page }))
    if (info.memberFilters?.length) pivot.memberFilters = info.memberFilters.map(({ field, excludedItems }) => ({ field, mode: 'exclude', values: [...excludedItems] }))
    if (info.sorts?.length) pivot.sorts = info.sorts.map((sort) => ({ ...sort }))
    pivots.push(pivot)
  }
  return { pivots, skipped }
}
