import type { OutlineAxis, OutlineGroup, OutlineResult } from './types'
import { MAX_SHEET_COLUMN_INDEX, MAX_SHEET_ROW_INDEX, validateOutlineSnapshot } from './validation'

/** SpreadsheetML stores outlineLevel as an unsigned value in the range 0..7. */
export const XLSX_MAX_OUTLINE_DEPTH = 7

export interface NativeXlsxOutlineBand {
  start: number
  end: number
  outlineLevel: number
  hidden: boolean
  collapsed: boolean
}

export interface NativeXlsxOutlineSnapshot {
  sheetId: string
  summaryBelow: boolean
  summaryRight: boolean
  rows: NativeXlsxOutlineBand[]
  columns: NativeXlsxOutlineBand[]
  groups: OutlineGroup[]
}

export interface NativeXlsxOutlineGroupWire {
  id: string
  sheet_id: string
  axis: OutlineAxis
  start: number
  end: number
  collapsed: boolean
}

export interface NativeXlsxOutlineBandWire {
  start: number
  end: number
  outline_level: number
  hidden: boolean
  collapsed: boolean
}

export interface NativeXlsxOutlineSnapshotWire {
  sheet_id: string
  summary_below: boolean
  summary_right: boolean
  rows: NativeXlsxOutlineBandWire[]
  columns: NativeXlsxOutlineBandWire[]
  groups: NativeXlsxOutlineGroupWire[]
}

export interface NativeXlsxOutlineWriteWire {
  sheet_id: string
  summary_below: boolean
  summary_right: boolean
  groups: NativeXlsxOutlineGroupWire[]
}

export type NativeXlsxOutlineDecodeResult =
  | { ok: true; value: NativeXlsxOutlineSnapshot }
  | { ok: false; issues: string[] }

const snapshotKeys = ['sheet_id', 'summary_below', 'summary_right', 'rows', 'columns', 'groups'] as const
const bandKeys = ['start', 'end', 'outline_level', 'hidden', 'collapsed'] as const
const groupKeys = ['id', 'sheet_id', 'axis', 'start', 'end', 'collapsed'] as const

function exactObject(value: unknown, keys: readonly string[], path: string, issues: string[]): Record<string, unknown> | null {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    issues.push(`${path} must be an object`)
    return null
  }
  const object = value as Record<string, unknown>
  const allowed = new Set(keys)
  for (const key of Object.keys(object)) {
    if (!allowed.has(key)) issues.push(`${path}.${key} is not supported`)
  }
  for (const key of keys) {
    if (!Object.hasOwn(object, key)) issues.push(`${path}.${key} is required`)
  }
  return object
}

function integer(value: unknown, minimum: number, maximum: number, path: string, issues: string[]): number {
  if (!Number.isSafeInteger(value) || (value as number) < minimum || (value as number) > maximum) {
    issues.push(`${path} must be an integer within ${minimum}-${maximum}`)
    return minimum
  }
  return value as number
}

function boolean(value: unknown, path: string, issues: string[]): boolean {
  if (typeof value !== 'boolean') {
    issues.push(`${path} must be a boolean`)
    return false
  }
  return value
}

function string(value: unknown, path: string, issues: string[]): string {
  if (typeof value !== 'string') {
    issues.push(`${path} must be a string`)
    return ''
  }
  return value
}

function decodeBands(value: unknown, maximum: number, path: string, issues: string[]): NativeXlsxOutlineBand[] {
  if (!Array.isArray(value)) {
    issues.push(`${path} must be an array`)
    return []
  }
  const bands: NativeXlsxOutlineBand[] = []
  let previousEnd = -1
  for (let index = 0; index < value.length; index += 1) {
    const itemPath = `${path}[${index}]`
    const object = exactObject(value[index], bandKeys, itemPath, issues)
    if (!object) continue
    const start = integer(object.start, 0, maximum, `${itemPath}.start`, issues)
    const end = integer(object.end, 0, maximum, `${itemPath}.end`, issues)
    const outlineLevel = integer(object.outline_level, 0, XLSX_MAX_OUTLINE_DEPTH, `${itemPath}.outline_level`, issues)
    const hidden = boolean(object.hidden, `${itemPath}.hidden`, issues)
    const collapsed = boolean(object.collapsed, `${itemPath}.collapsed`, issues)
    if (end < start) issues.push(`${itemPath} has reversed bounds`)
    if (start <= previousEnd) issues.push(`${itemPath} overlaps or is not strictly ordered`)
    previousEnd = Math.max(previousEnd, end)
    bands.push({ start, end, outlineLevel, hidden, collapsed })
  }
  return bands
}

/** Decode the exact JSON shape emitted by Go's ReadWorksheetOutline. */
export function decodeNativeXlsxOutlineSnapshot(value: unknown): NativeXlsxOutlineDecodeResult {
  const issues: string[] = []
  const object = exactObject(value, snapshotKeys, '$', issues)
  if (!object) return { ok: false, issues }
  const sheetId = string(object.sheet_id, '$.sheet_id', issues)
  if (sheetId.length === 0 || sheetId.length > 128 || sheetId.trim() !== sheetId) {
    issues.push('$.sheet_id must contain 1-128 unpadded characters')
  }
  const summaryBelow = boolean(object.summary_below, '$.summary_below', issues)
  const summaryRight = boolean(object.summary_right, '$.summary_right', issues)
  const rows = decodeBands(object.rows, MAX_SHEET_ROW_INDEX, '$.rows', issues)
  const columns = decodeBands(object.columns, MAX_SHEET_COLUMN_INDEX, '$.columns', issues)
  const groups: OutlineGroup[] = []
  if (!Array.isArray(object.groups)) {
    issues.push('$.groups must be an array')
  } else {
    for (let index = 0; index < object.groups.length; index += 1) {
      const path = `$.groups[${index}]`
      const group = exactObject(object.groups[index], groupKeys, path, issues)
      if (!group) continue
      const axis = group.axis
      if (axis !== 'row' && axis !== 'column') issues.push(`${path}.axis must be row or column`)
      const maximum = axis === 'column' ? MAX_SHEET_COLUMN_INDEX : MAX_SHEET_ROW_INDEX
      groups.push({
        id: string(group.id, `${path}.id`, issues),
        sheetId: string(group.sheet_id, `${path}.sheet_id`, issues),
        axis: axis === 'column' ? 'column' : 'row',
        start: integer(group.start, 0, maximum, `${path}.start`, issues),
        end: integer(group.end, 0, maximum, `${path}.end`, issues),
        collapsed: boolean(group.collapsed, `${path}.collapsed`, issues),
      })
    }
  }
  const validated = validateOutlineSnapshot(groups, XLSX_MAX_OUTLINE_DEPTH)
  if (!validated.ok) issues.push(...validated.issues.map((issue) => `$.groups: ${issue.message}`))
  for (const [index, group] of groups.entries()) {
    if (group.sheetId !== sheetId) issues.push(`$.groups[${index}].sheet_id must match $.sheet_id`)
  }
  if (issues.length > 0) return { ok: false, issues }
  return { ok: true, value: { sheetId, summaryBelow, summaryRight, rows, columns, groups: validated.ok ? validated.value : [] } }
}

/** Build the exact JSON shape accepted by Go's ApplyWorksheetOutline. */
export function createNativeXlsxOutlineWrite(
  sheetId: string,
  groups: readonly OutlineGroup[],
  options: { summaryBelow?: boolean; summaryRight?: boolean } = {},
): OutlineResult<NativeXlsxOutlineWriteWire> {
  const validated = validateOutlineSnapshot(groups, XLSX_MAX_OUTLINE_DEPTH)
  if (!validated.ok) return validated
  if (sheetId.length === 0 || sheetId.length > 128 || sheetId.trim() !== sheetId || validated.value.some((group) => group.sheetId !== sheetId)) {
    return { ok: false, issues: [{ code: 'INVALID_SHEET', message: 'sheetId must match every group and contain 1-128 unpadded characters' }] }
  }
  return {
    ok: true,
    value: {
      sheet_id: sheetId,
      summary_below: options.summaryBelow ?? true,
      summary_right: options.summaryRight ?? true,
      groups: validated.value.map((group) => ({
        id: group.id,
        sheet_id: group.sheetId,
        axis: group.axis,
        start: group.start,
        end: group.end,
        collapsed: group.collapsed,
      })),
    },
  }
}
