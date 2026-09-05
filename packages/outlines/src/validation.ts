import type { OutlineGroup, OutlineIssue, OutlineResult } from './types'

export const MAX_OUTLINE_DEPTH = 8
export const MAX_SHEET_ROW_INDEX = 1_048_575
export const MAX_SHEET_COLUMN_INDEX = 16_383

function contains(outer: OutlineGroup, inner: OutlineGroup): boolean {
  return outer.start <= inner.start && outer.end >= inner.end
}

function overlaps(left: OutlineGroup, right: OutlineGroup): boolean {
  return left.start <= right.end && right.start <= left.end
}

function sameLane(left: OutlineGroup, right: OutlineGroup): boolean {
  return left.sheetId === right.sheetId && left.axis === right.axis
}

export function validateOutlineGroup(
  input: OutlineGroup,
  existing: readonly OutlineGroup[] = [],
  maxDepth = MAX_OUTLINE_DEPTH,
): OutlineResult<OutlineGroup> {
  const issues: OutlineIssue[] = []
  if (!/^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/.test(input.id)) {
    issues.push({ code: 'INVALID_ID', message: 'id must use 1-128 safe ASCII identifier characters' })
  }
  if (input.sheetId.length === 0 || input.sheetId.length > 128) {
    issues.push({ code: 'INVALID_SHEET', message: 'sheetId must contain 1-128 characters' })
  }
  if (input.axis !== 'row' && input.axis !== 'column') {
    issues.push({ code: 'INVALID_AXIS', message: 'axis must be row or column' })
  }
  if (typeof input.collapsed !== 'boolean') {
    issues.push({ code: 'INVALID_COLLAPSED', message: 'collapsed must be a boolean' })
  }
  const maximum = input.axis === 'column' ? MAX_SHEET_COLUMN_INDEX : MAX_SHEET_ROW_INDEX
  if (!Number.isSafeInteger(input.start) || !Number.isSafeInteger(input.end) || input.start < 0 || input.end < input.start || input.end > maximum) {
    issues.push({ code: 'INVALID_RANGE', message: `range must be an inclusive interval within 0-${maximum}` })
  }

  for (const group of existing) {
    if (group.id === input.id) {
      issues.push({ code: 'DUPLICATE', message: `outline id ${input.id} already exists`, groupId: group.id })
      continue
    }
    if (!sameLane(group, input)) continue
    if (group.start === input.start && group.end === input.end) {
      issues.push({ code: 'DUPLICATE', message: 'an outline already has the same range', groupId: group.id })
    } else if (overlaps(group, input) && !contains(group, input) && !contains(input, group)) {
      issues.push({ code: 'CROSSING', message: 'outline ranges may be disjoint or nested, but cannot cross', groupId: group.id })
    }
  }

  if (issues.length === 0) {
    const all = [...existing, input]
    for (const group of all) {
      const depth = all.filter((candidate) =>
        candidate !== group && sameLane(candidate, group) && contains(candidate, group),
      ).length + 1
      if (depth > maxDepth) {
        issues.push({ code: 'MAX_DEPTH', message: `outline nesting cannot exceed ${maxDepth} levels`, groupId: group.id })
        break
      }
    }
  }

  return issues.length === 0 ? { ok: true, value: Object.freeze({ ...input }) } : { ok: false, issues }
}

export function validateOutlineSnapshot(groups: readonly OutlineGroup[], maxDepth = MAX_OUTLINE_DEPTH): OutlineResult<OutlineGroup[]> {
  const accepted: OutlineGroup[] = []
  const issues: OutlineIssue[] = []
  for (const group of groups) {
    const result = validateOutlineGroup(group, accepted, maxDepth)
    if (result.ok) accepted.push(result.value)
    else issues.push(...result.issues)
  }
  return issues.length === 0 ? { ok: true, value: accepted } : { ok: false, issues }
}
