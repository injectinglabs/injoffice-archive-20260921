import { SPARKLINE_TYPES } from './types'
import type { SparklineOptions, SparklineSnapshotV1, SparklineSpec, SparklineType } from './types'

/** JSON shape returned by go/xlsxpatch.ReadSparklines. */
export interface FileSparklineInfo {
  id: string
  groupId?: string
  type: string
  sourceSheetName?: string
  sourceRef?: string
  targetSheetName: string
  targetRef: string
  options?: SparklineOptions
  warnings?: string[]
}

export interface SparklineFileContext {
  sheetIdOf: (sheetName: string) => string | null
}

export interface SparklineFileResult {
  snapshot: SparklineSnapshotV1
  skipped: string[]
}

const rangePattern = /^\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/i
const cellPattern = /^\$?([A-Z]{1,3})\$?([1-9]\d*)$/i

function columnIndex(label: string): number {
  let value = 0
  for (const char of label.toUpperCase()) value = value * 26 + char.charCodeAt(0) - 64
  return value - 1
}

function parseSource(value: string, sheetId: string): SparklineSpec['source'] | null {
  const match = rangePattern.exec(value.trim())
  if (!match) return null
  const startRow = Number(match[2]) - 1
  const startColumn = columnIndex(match[1])
  const endRow = Number(match[4] ?? match[2]) - 1
  const endColumn = columnIndex(match[3] ?? match[1])
  if (endRow > 1_048_575 || endColumn > 16_383 || endRow < startRow || endColumn < startColumn || (endRow > startRow && endColumn > startColumn)) return null
  return { sheetId, startRow, startColumn, endRow, endColumn }
}

function parseTarget(value: string, sheetId: string): SparklineSpec['target'] | null {
  const match = cellPattern.exec(value.trim())
  if (!match) return null
  const row = Number(match[2]) - 1
  const column = columnIndex(match[1])
  return row <= 1_048_575 && column <= 16_383 ? { sheetId, row, column } : null
}

const safeId = (value: string): string => value.replace(/[^A-Za-z0-9_-]+/g, '_').replace(/^_+|_+$/g, '') || 'sparkline'

/**
 * Converts the native reader output atomically per native group. A warning or
 * unavailable sheet skips the whole group, because retaining only some members
 * would silently change shared axis and styling semantics.
 */
export function sparklinesFromFile(infos: FileSparklineInfo[], context: SparklineFileContext): SparklineFileResult {
  const sparklines: SparklineSpec[] = []
  const groups: SparklineSnapshotV1['groups'] = []
  const skipped: string[] = []
  const buckets = new Map<string, FileSparklineInfo[]>()
  infos.forEach((info, index) => {
    const key = info.groupId ? `group:${info.groupId}` : `item:${index}`
    buckets.set(key, [...(buckets.get(key) ?? []), info])
  })
  const usedIds = new Set<string>()
  for (const [bucketKey, members] of buckets) {
    const nativeLabel = members[0].groupId ?? members[0].id
    const reason = members.flatMap((member) => member.warnings ?? [])
    if (reason.length) {
      skipped.push(`sparkline ${nativeLabel}: ${[...new Set(reason)].join('; ')}`)
      continue
    }
    const converted: SparklineSpec[] = []
    let failure = ''
    for (const member of members) {
      if (!(SPARKLINE_TYPES as readonly string[]).includes(member.type)) { failure = `unsupported type ${member.type}`; break }
      if (!member.sourceSheetName || !member.sourceRef) { failure = 'source reference is missing'; break }
      const sourceSheetId = context.sheetIdOf(member.sourceSheetName)
      const targetSheetId = context.sheetIdOf(member.targetSheetName)
      if (!sourceSheetId || !targetSheetId) { failure = 'source or target sheet is unavailable'; break }
      const source = parseSource(member.sourceRef, sourceSheetId)
      const target = parseTarget(member.targetRef, targetSheetId)
      if (!source || !target) { failure = 'source or target A1 reference is unsupported'; break }
      let id = safeId(member.id)
      for (let suffix = 2; usedIds.has(id); suffix++) id = `${safeId(member.id)}_${suffix}`
      usedIds.add(id)
      converted.push({ id, type: member.type as SparklineType, source, target, options: member.options })
    }
    if (failure) {
      skipped.push(`sparkline ${nativeLabel}: ${failure}`)
      for (const spec of converted) usedIds.delete(spec.id)
      continue
    }
    if (bucketKey.startsWith('group:') && converted.length > 1) {
      let groupId = safeId(nativeLabel)
      for (let suffix = 2; groups.some((group) => group.id === groupId); suffix++) groupId = `${safeId(nativeLabel)}_${suffix}`
      for (const spec of converted) spec.groupId = groupId
      groups.push({ id: groupId, memberIds: converted.map((spec) => spec.id) })
    }
    sparklines.push(...converted)
  }
  return { snapshot: { version: 1, sparklines, groups }, skipped }
}
