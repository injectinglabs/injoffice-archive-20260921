import type { SparklineOptions, SparklineRangeRef, SparklineSnapshotV1, SparklineSpec } from './types'
import { validateSparklineSnapshot } from './validation'

export interface WireSparklineItem {
  sourceSheetName: string
  sourceRef: string
  targetCellRef: string
}

export interface WireSparklineGroup {
  id?: string
  targetSheetName: string
  type: SparklineSpec['type']
  options?: SparklineOptions
  sparklines: WireSparklineItem[]
}

export interface SparklineWireContext {
  sheetNameOf: (sheetId: string) => string | null
}

export interface SparklineWireResult {
  groups: WireSparklineGroup[]
  skipped: string[]
}

function columnName(column: number): string {
  let value = column + 1
  let label = ''
  while (value > 0) {
    label = String.fromCharCode(65 + (value - 1) % 26) + label
    value = Math.floor((value - 1) / 26)
  }
  return label
}

const cellA1 = (row: number, column: number): string => `${columnName(column)}${row + 1}`
const rangeA1 = (range: SparklineRangeRef): string => `${cellA1(range.startRow, range.startColumn)}:${cellA1(range.endRow, range.endColumn)}`

function optionsKey(options: SparklineOptions | undefined): string {
  const colors = options?.colors
  return JSON.stringify({
    emptyCells: options?.emptyCells ?? 'gap', rightToLeft: options?.rightToLeft ?? false,
    showMarkers: options?.showMarkers ?? false, showHigh: options?.showHigh ?? false,
    showLow: options?.showLow ?? false, showFirst: options?.showFirst ?? false,
    showLast: options?.showLast ?? false, showNegative: options?.showNegative ?? false,
    lineWeight: options?.lineWeight, min: options?.min, max: options?.max,
    colors: colors ? { series: colors.series, negative: colors.negative, markers: colors.markers, high: colors.high, low: colors.low, first: colors.first, last: colors.last, axis: colors.axis } : undefined,
  })
}

function unsupportedColor(options: SparklineOptions | undefined): string | null {
  for (const value of Object.values(options?.colors ?? {})) {
    if (value !== undefined && !/^#?(?:[0-9a-f]{6}|FF[0-9a-f]{6})$/i.test(value)) return value
  }
  return null
}

/**
 * Converts a complete snapshot to native x14 groups. Conversion is fail-closed
 * per model group: cross-sheet targets, unequal shared options, missing sheets,
 * and colors outside the native RGB subset are reported instead of guessed.
 */
export function toWireSparklines(snapshot: SparklineSnapshotV1, context: SparklineWireContext): SparklineWireResult {
  const issues = validateSparklineSnapshot(snapshot)
  if (issues.length) return { groups: [], skipped: [`sparkline snapshot: ${issues.map((issue) => `${issue.path} ${issue.message}`).join('; ')}`] }
  const byId = new Map(snapshot.sparklines.map((spec) => [spec.id, spec]))
  const buckets: Array<{ id?: string; members: SparklineSpec[] }> = snapshot.groups.map((group) => ({ id: group.id, members: group.memberIds.map((id) => byId.get(id)!) }))
  for (const spec of snapshot.sparklines) if (!spec.groupId) buckets.push({ members: [spec] })
  const groups: WireSparklineGroup[] = []
  const skipped: string[] = []
  for (const bucket of buckets) {
    const label = bucket.id ?? bucket.members[0].id
    const targetNames = bucket.members.map((spec) => context.sheetNameOf(spec.target.sheetId))
    const sourceNames = bucket.members.map((spec) => context.sheetNameOf(spec.source.sheetId))
    const exceedsExcelBounds = bucket.members.some((spec) => spec.source.endRow > 1_048_575 || spec.source.endColumn > 16_383 || spec.target.row > 1_048_575 || spec.target.column > 16_383)
    if (exceedsExcelBounds) {
      skipped.push(`sparkline ${label}: its source or target exceeds native XLSX bounds`)
      continue
    }
    if (targetNames.some((name) => !name) || sourceNames.some((name) => !name)) {
      skipped.push(`sparkline ${label}: its source or target sheet no longer exists`)
      continue
    }
    if (new Set(targetNames).size !== 1) {
      skipped.push(`sparkline group ${label}: native group targets must be on one worksheet`)
      continue
    }
    if (new Set(bucket.members.map((spec) => spec.type)).size !== 1) {
      skipped.push(`sparkline group ${label}: members must have one type`)
      continue
    }
    if (new Set(bucket.members.map((spec) => optionsKey(spec.options))).size !== 1) {
      skipped.push(`sparkline group ${label}: native group members must share identical options`)
      continue
    }
    const badColor = unsupportedColor(bucket.members[0].options)
    if (badColor) {
      skipped.push(`sparkline ${label}: color ${badColor} is not a supported native RGB value`)
      continue
    }
    groups.push({
      id: bucket.id,
      targetSheetName: targetNames[0]!,
      type: bucket.members[0].type,
      options: bucket.members[0].options,
      sparklines: bucket.members.map((spec, index) => ({
        sourceSheetName: sourceNames[index]!, sourceRef: rangeA1(spec.source), targetCellRef: cellA1(spec.target.row, spec.target.column),
      })),
    })
  }
  return { groups, skipped }
}
