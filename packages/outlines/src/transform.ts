import type { OutlineGroup, OutlineStructuralEdit } from './types'

/** Transform one outline through a structural edit ordered before it. */
export function transformOutline(group: OutlineGroup, edit: OutlineStructuralEdit): OutlineGroup | null {
  if (group.sheetId !== edit.sheetId || group.axis !== edit.axis || edit.count <= 0) return group
  if (edit.kind === 'insert') {
    if (edit.start <= group.start) return { ...group, start: group.start + edit.count, end: group.end + edit.count }
    if (edit.start <= group.end) return { ...group, end: group.end + edit.count }
    return group
  }

  const removedEnd = edit.start + edit.count - 1
  if (edit.start > group.end) return group
  if (removedEnd < group.start) return { ...group, start: group.start - edit.count, end: group.end - edit.count }
  if (edit.start <= group.start && removedEnd >= group.end) return null
  if (edit.start <= group.start) return { ...group, start: edit.start, end: group.end - edit.count }
  if (removedEnd <= group.end) return { ...group, end: group.end - edit.count }
  return { ...group, end: edit.start - 1 }
}

export function transformOutlines(groups: readonly OutlineGroup[], edit: OutlineStructuralEdit): OutlineGroup[] {
  return groups.map((group) => transformOutline(group, edit)).filter((group): group is OutlineGroup => group !== null)
}
