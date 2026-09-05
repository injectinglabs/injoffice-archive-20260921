import type { OutlineAxis, OutlineGroup, OutlineIssue, OutlineResult } from './types'
import { validateOutlineGroup, validateOutlineSnapshot } from './validation'

function ordered(groups: Iterable<OutlineGroup>): OutlineGroup[] {
  return [...groups].sort((left, right) =>
    left.sheetId.localeCompare(right.sheetId) || left.axis.localeCompare(right.axis) || left.start - right.start || right.end - left.end || left.id.localeCompare(right.id),
  )
}

export class OutlineStore {
  private groups = new Map<string, OutlineGroup>()

  add(group: OutlineGroup): OutlineResult<OutlineGroup> {
    const result = validateOutlineGroup(group, this.list())
    if (!result.ok) return result
    this.groups.set(result.value.id, result.value)
    return result
  }

  update(id: string, patch: Partial<Omit<OutlineGroup, 'id'>>): OutlineResult<OutlineGroup> {
    const current = this.groups.get(id)
    if (!current) return { ok: false, issues: [{ code: 'INVALID_ID', message: `outline ${id} does not exist` }] }
    const candidate = { ...current, ...patch, id }
    const result = validateOutlineGroup(candidate, this.list().filter((group) => group.id !== id))
    if (!result.ok) return result
    this.groups.set(id, result.value)
    return result
  }

  remove(id: string): OutlineGroup | undefined {
    const group = this.groups.get(id)
    if (group) this.groups.delete(id)
    return group
  }

  get(id: string): OutlineGroup | undefined {
    return this.groups.get(id)
  }

  list(sheetId?: string, axis?: OutlineAxis): OutlineGroup[] {
    return ordered([...this.groups.values()].filter((group) =>
      (sheetId === undefined || group.sheetId === sheetId) && (axis === undefined || group.axis === axis),
    ))
  }

  clear(sheetId?: string, axis?: OutlineAxis, start?: number, end?: number): OutlineGroup[] {
    const removed: OutlineGroup[] = []
    for (const group of this.list(sheetId, axis)) {
      if (start !== undefined && group.start < start) continue
      if (end !== undefined && group.end > end) continue
      this.groups.delete(group.id)
      removed.push(group)
    }
    return removed
  }

  hydrate(groups: readonly OutlineGroup[]): OutlineResult<OutlineGroup[]> {
    const result = validateOutlineSnapshot(groups)
    if (!result.ok) return result
    this.groups = new Map(result.value.map((group) => [group.id, group]))
    return { ok: true, value: this.list() }
  }

  depth(id: string): number | undefined {
    const group = this.groups.get(id)
    if (!group) return undefined
    return this.list(group.sheetId, group.axis).filter((candidate) =>
      candidate.id !== id && candidate.start <= group.start && candidate.end >= group.end,
    ).length + 1
  }
}

export function outlineFailure(message: string): OutlineResult<never> {
  const issue: OutlineIssue = { code: 'INVALID_ID', message }
  return { ok: false, issues: [issue] }
}
