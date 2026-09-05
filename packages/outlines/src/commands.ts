import type { OutlineAxis, OutlineGroup, OutlineResult } from './types'
import { OUTLINE_AUTHORITATIVE_APPLY, OutlineManager } from './manager'

export interface OutlineUndoRecord {
  label: string
  before: OutlineGroup[]
  after: OutlineGroup[]
}

export interface OutlineUndoSink {
  push(record: OutlineUndoRecord): void
}

export interface OutlineAuthoritativeApplyOptions {
  label?: string
  recordUndo?: boolean
}

/** @internal Capability used only by the ordered collaboration session. */
export const OUTLINE_COLLABORATION_APPLY = Symbol('injoffice.outline.collaboration-apply')

export class OutlineHandle {
  constructor(private readonly owner: OutlineCommandController, readonly id: string) {}

  get value(): OutlineGroup | undefined {
    const group = this.owner.manager.get(this.id)
    return group ? { ...group } : undefined
  }

  update(patch: Partial<Omit<OutlineGroup, 'id'>>): OutlineResult<OutlineGroup> {
    return this.owner.update(this.id, patch)
  }

  setCollapsed(collapsed: boolean): OutlineResult<OutlineGroup> {
    return this.owner.setCollapsed(this.id, collapsed)
  }

  toggle(): OutlineResult<OutlineGroup> {
    const value = this.value
    return value ? this.owner.setCollapsed(this.id, !value.collapsed) : this.owner.manager.toggle(this.id)
  }

  remove(): boolean {
    return this.owner.remove(this.id)
  }
}

/** Host-neutral lifecycle facade with atomic before/after undo records. */
export class OutlineCommandController {
  constructor(readonly manager: OutlineManager, private readonly undo?: OutlineUndoSink) {}

  getById(id: string): OutlineHandle | null {
    return this.manager.get(id) ? new OutlineHandle(this, id) : null
  }

  list(sheetId?: string, axis?: OutlineAxis): OutlineHandle[] {
    return this.manager.list(sheetId, axis).map((group) => new OutlineHandle(this, group.id))
  }

  snapshot(): OutlineGroup[] {
    return this.manager.serialize()
  }

  add(group: OutlineGroup): OutlineResult<OutlineGroup> {
    return this.record('Add outline', () => this.manager.add(group))
  }

  update(id: string, patch: Partial<Omit<OutlineGroup, 'id'>>): OutlineResult<OutlineGroup> {
    return this.record('Update outline', () => this.manager.update(id, patch))
  }

  setCollapsed(id: string, collapsed: boolean): OutlineResult<OutlineGroup> {
    return this.record(collapsed ? 'Collapse outline' : 'Expand outline', () => this.manager.setCollapsed(id, collapsed))
  }

  remove(id: string): boolean {
    return this.record('Remove outline', () => this.manager.remove(id))
  }

  clear(sheetId: string, axis?: OutlineAxis, start?: number, end?: number): OutlineGroup[] {
    return this.record('Ungroup outline', () => this.manager.clear(sheetId, axis, start, end))
  }

  restore(snapshot: readonly OutlineGroup[]): boolean {
    return this.manager.hydrate(snapshot).ok
  }

  /** Apply one already-authorized complete snapshot. Local acknowledged
   * commands may record undo; remote entries and resync omit local history. */
  [OUTLINE_COLLABORATION_APPLY](snapshot: readonly OutlineGroup[], options: OutlineAuthoritativeApplyOptions = {}): boolean {
    const before = this.snapshot()
    const result = this.manager[OUTLINE_AUTHORITATIVE_APPLY](snapshot)
    if (!result.ok) return false
    const after = this.snapshot()
    if (options.recordUndo === true && JSON.stringify(before) !== JSON.stringify(after)) {
      const label = typeof options.label === 'string' && options.label.trim() && options.label.length <= 256
        ? options.label
        : 'Apply outline snapshot'
      this.undo?.push({ label, before, after })
    }
    return true
  }

  private record<T>(label: string, mutation: () => T): T {
    const before = this.manager.serialize()
    const value = mutation()
    const after = this.manager.serialize()
    if (JSON.stringify(before) !== JSON.stringify(after)) this.undo?.push({ label, before, after })
    return value
  }
}
