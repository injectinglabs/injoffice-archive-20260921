import { SPARKLINE_TYPES } from './types'
import type { SparklineSnapshotV1, SparklineSpec } from './types'

export interface SparklineValidationIssue {
  path: string
  message: string
}

const isIndex = (value: number): boolean => Number.isInteger(value) && value >= 0

export function validateSparkline(spec: SparklineSpec): SparklineValidationIssue[] {
  const issues: SparklineValidationIssue[] = []
  if (!spec.id.trim()) issues.push({ path: 'id', message: 'must not be empty' })
  if (!SPARKLINE_TYPES.includes(spec.type)) issues.push({ path: 'type', message: 'must be line, column, or win-loss' })
  if (!spec.source.sheetId.trim()) issues.push({ path: 'source.sheetId', message: 'must not be empty' })
  if (!spec.target.sheetId.trim()) issues.push({ path: 'target.sheetId', message: 'must not be empty' })

  for (const key of ['startRow', 'startColumn', 'endRow', 'endColumn'] as const) {
    if (!isIndex(spec.source[key])) issues.push({ path: `source.${key}`, message: 'must be a non-negative integer' })
  }
  for (const key of ['row', 'column'] as const) {
    if (!isIndex(spec.target[key])) issues.push({ path: `target.${key}`, message: 'must be a non-negative integer' })
  }
  if (spec.source.endRow < spec.source.startRow) issues.push({ path: 'source.endRow', message: 'must not precede startRow' })
  if (spec.source.endColumn < spec.source.startColumn) issues.push({ path: 'source.endColumn', message: 'must not precede startColumn' })
  if (spec.source.endRow > spec.source.startRow && spec.source.endColumn > spec.source.startColumn) {
    issues.push({ path: 'source', message: 'must contain exactly one row or one column' })
  }
  if (spec.groupId !== undefined && !spec.groupId.trim()) issues.push({ path: 'groupId', message: 'must not be empty' })
  if (
    spec.source.sheetId === spec.target.sheetId &&
    spec.target.row >= spec.source.startRow && spec.target.row <= spec.source.endRow &&
    spec.target.column >= spec.source.startColumn && spec.target.column <= spec.source.endColumn
  ) issues.push({ path: 'target', message: 'must not overlap the source range' })

  const options = spec.options
  if (options?.lineWeight !== undefined && (!Number.isFinite(options.lineWeight) || options.lineWeight <= 0)) {
    issues.push({ path: 'options.lineWeight', message: 'must be a finite positive number' })
  }
  for (const key of ['min', 'max'] as const) {
    const value = options?.[key]
    if (value !== undefined && !Number.isFinite(value)) issues.push({ path: `options.${key}`, message: 'must be finite' })
  }
  if (options?.min !== undefined && options.max !== undefined && options.min >= options.max) {
    issues.push({ path: 'options', message: 'min must be less than max' })
  }
  return issues
}

export function validateSparklineSnapshot(snapshot: SparklineSnapshotV1): SparklineValidationIssue[] {
  const issues: SparklineValidationIssue[] = []
  if (snapshot.version !== 1) issues.push({ path: 'version', message: 'must be 1' })
  const ids = new Set<string>()
  const targets = new Set<string>()
  snapshot.sparklines.forEach((spec, index) => {
    for (const issue of validateSparkline(spec)) issues.push({ path: `sparklines.${index}.${issue.path}`, message: issue.message })
    if (ids.has(spec.id)) issues.push({ path: `sparklines.${index}.id`, message: 'must be unique' })
    ids.add(spec.id)
    const target = `${spec.target.sheetId}\u0000${spec.target.row}\u0000${spec.target.column}`
    if (targets.has(target)) issues.push({ path: `sparklines.${index}.target`, message: 'must be unique' })
    targets.add(target)
  })
  const groupIds = new Set<string>()
  snapshot.groups.forEach((group, index) => {
    if (!group.id.trim()) issues.push({ path: `groups.${index}.id`, message: 'must not be empty' })
    if (groupIds.has(group.id)) issues.push({ path: `groups.${index}.id`, message: 'must be unique' })
    groupIds.add(group.id)
    if (group.memberIds.length < 2) issues.push({ path: `groups.${index}.memberIds`, message: 'must contain at least two members' })
    const members = new Set<string>()
    let memberType: SparklineSpec['type'] | undefined
    for (const memberId of group.memberIds) {
      if (!ids.has(memberId)) issues.push({ path: `groups.${index}.memberIds`, message: `unknown sparkline ${memberId}` })
      if (members.has(memberId)) issues.push({ path: `groups.${index}.memberIds`, message: `duplicate sparkline ${memberId}` })
      members.add(memberId)
      const spec = snapshot.sparklines.find((candidate) => candidate.id === memberId)
      if (spec?.groupId !== group.id) issues.push({ path: `sparklines.${memberId}.groupId`, message: `must equal ${group.id}` })
      if (spec && memberType !== undefined && spec.type !== memberType) issues.push({ path: `groups.${index}.memberIds`, message: 'members must have the same type' })
      if (spec && memberType === undefined) memberType = spec.type
    }
  })
  for (const spec of snapshot.sparklines) {
    if (spec.groupId && !groupIds.has(spec.groupId)) issues.push({ path: `sparklines.${spec.id}.groupId`, message: `unknown group ${spec.groupId}` })
  }
  return issues
}
