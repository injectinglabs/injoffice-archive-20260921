import { OutlineStore, outlineFailure } from './store'
import type { OutlineAxis, OutlineGroup, OutlineManagerOptions, OutlineResult, OutlineVisibilityAdapter } from './types'
import { validateOutlineSnapshot } from './validation'

export const OUTLINE_AUTHORITATIVE_APPLY = Symbol('injoffice.outline.authoritative-apply')

type Interval = readonly [start: number, end: number]

function mergedCoverage(groups: readonly OutlineGroup[], start: number, end: number): Interval[] {
  const spans = groups
    .filter((group) => group.collapsed && group.start <= end && group.end >= start)
    .map((group) => [Math.max(start, group.start), Math.min(end, group.end)] as Interval)
    .sort((left, right) => left[0] - right[0] || left[1] - right[1])
  const merged: Array<[number, number]> = []
  for (const span of spans) {
    const previous = merged.at(-1)
    if (!previous || span[0] > previous[1] + 1) merged.push([span[0], span[1]])
    else previous[1] = Math.max(previous[1], span[1])
  }
  return merged
}

/** Owns outline visibility for the ranges registered through this manager. */
export class OutlineManager {
  private readonly store: OutlineStore
  private readonly visibility: OutlineVisibilityAdapter
  private readonly listeners = new Set<() => void>()
  readonly mutationAuthority: 'local' | 'collaboration'

  constructor(visibility: OutlineVisibilityAdapter, options?: OutlineManagerOptions)
  constructor(visibility: OutlineVisibilityAdapter, store?: OutlineStore, options?: OutlineManagerOptions)
  constructor(
    visibility: OutlineVisibilityAdapter,
    storeOrOptions: OutlineStore | OutlineManagerOptions = new OutlineStore(),
    explicitOptions: OutlineManagerOptions = {},
  ) {
    const store = storeOrOptions instanceof OutlineStore ? storeOrOptions : new OutlineStore()
    const options = storeOrOptions instanceof OutlineStore ? explicitOptions : storeOrOptions
    this.visibility = visibility
    this.mutationAuthority = options.mutationAuthority ?? 'local'
    if (this.mutationAuthority === 'collaboration') {
      this.store = new OutlineStore()
      this.store.hydrate(store.list())
    } else this.store = store
  }

  onChange(listener: () => void): () => void {
    this.listeners.add(listener)
    return () => this.listeners.delete(listener)
  }

  add(group: OutlineGroup): OutlineResult<OutlineGroup> {
    if (this.mutationAuthority === 'collaboration') return authorityFailure()
    const result = this.store.add(group)
    if (!result.ok) return result
    if (group.collapsed) this.visibility.hide(group.sheetId, group.axis, group.start, group.end - group.start + 1)
    this.emit()
    return result
  }

  update(id: string, patch: Partial<Omit<OutlineGroup, 'id'>>): OutlineResult<OutlineGroup> {
    if (this.mutationAuthority === 'collaboration') return authorityFailure()
    const previous = this.store.get(id)
    if (!previous) return outlineFailure(`outline ${id} does not exist`)
    const result = this.store.update(id, patch)
    if (!result.ok) return result
    this.reconcile(previous.sheetId, previous.axis, previous.start, previous.end)
    this.reconcile(result.value.sheetId, result.value.axis, result.value.start, result.value.end)
    this.emit()
    return result
  }

  setCollapsed(id: string, collapsed: boolean): OutlineResult<OutlineGroup> {
    return this.update(id, { collapsed })
  }

  toggle(id: string): OutlineResult<OutlineGroup> {
    const group = this.store.get(id)
    return group ? this.setCollapsed(id, !group.collapsed) : outlineFailure(`outline ${id} does not exist`)
  }

  remove(id: string): boolean {
    if (this.mutationAuthority === 'collaboration') return false
    const group = this.store.remove(id)
    if (!group) return false
    this.reconcile(group.sheetId, group.axis, group.start, group.end)
    this.emit()
    return true
  }

  clear(sheetId: string, axis?: OutlineAxis, start?: number, end?: number): OutlineGroup[] {
    if (this.mutationAuthority === 'collaboration') return []
    const removed = this.store.clear(sheetId, axis, start, end)
    for (const group of removed) this.reconcile(group.sheetId, group.axis, group.start, group.end)
    if (removed.length > 0) this.emit()
    return removed
  }

  list(sheetId?: string, axis?: OutlineAxis): OutlineGroup[] {
    return this.store.list(sheetId, axis)
  }

  get(id: string): OutlineGroup | undefined {
    const group = this.store.get(id)
    return group ? { ...group } : undefined
  }

  depth(id: string): number | undefined {
    return this.store.depth(id)
  }

  /** Return a detached snapshot suitable for persistence or an undo record. */
  serialize(): OutlineGroup[] {
    return this.store.list().map((group) => ({ ...group }))
  }

  /**
   * Replace all groups as one model mutation. Validation happens before the
   * current store is touched, so malformed undo/import payloads fail closed.
   */
  hydrate(groups: readonly OutlineGroup[]): OutlineResult<OutlineGroup[]> {
    if (this.mutationAuthority === 'collaboration') return authorityFailure()
    return this.replace(groups)
  }

  /** @internal Collaboration sessions use this only after an authoritative
   * ordered acknowledgement or room-bound resync. */
  [OUTLINE_AUTHORITATIVE_APPLY](groups: readonly OutlineGroup[]): OutlineResult<OutlineGroup[]> {
    return this.replace(groups)
  }

  private replace(groups: readonly OutlineGroup[]): OutlineResult<OutlineGroup[]> {
    const previous = this.serialize()
    const checked = validateOutlineSnapshot(groups)
    if (!checked.ok) return checked
    const result = this.store.hydrate(checked.value)
    if (!result.ok) return result

    const affected = new Map<string, { sheetId: string; axis: OutlineAxis; start: number; end: number }>()
    for (const group of [...previous, ...result.value]) {
      const key = `${group.sheetId}\0${group.axis}`
      const extent = affected.get(key)
      if (extent) {
        extent.start = Math.min(extent.start, group.start)
        extent.end = Math.max(extent.end, group.end)
      } else {
        affected.set(key, { sheetId: group.sheetId, axis: group.axis, start: group.start, end: group.end })
      }
    }
    try {
      for (const extent of affected.values()) this.reconcile(extent.sheetId, extent.axis, extent.start, extent.end)
    } catch {
      this.store.hydrate(previous)
      try { for (const extent of affected.values()) this.reconcile(extent.sheetId, extent.axis, extent.start, extent.end) } catch { /* best-effort visibility rollback */ }
      return { ok: false, issues: [{ code: 'APPLY_FAILED', message: 'outline visibility snapshot could not be applied atomically' }] }
    }
    this.emit()
    return { ok: true, value: this.serialize() }
  }

  private reconcile(sheetId: string, axis: OutlineAxis, start: number, end: number): void {
    this.visibility.show(sheetId, axis, start, end - start + 1)
    for (const [coveredStart, coveredEnd] of mergedCoverage(this.store.list(sheetId, axis), start, end)) {
      this.visibility.hide(sheetId, axis, coveredStart, coveredEnd - coveredStart + 1)
    }
  }

  private emit(): void {
    for (const listener of this.listeners) { try { listener() } catch { /* observer isolation */ } }
  }
}

function authorityFailure<T>(): OutlineResult<T> {
  return { ok: false, issues: [{ code: 'AUTHORITY_REQUIRED', message: 'outline mutations require an acknowledged collaboration operation' }] }
}
