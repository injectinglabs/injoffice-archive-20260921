import type { FileChartAnchor } from './fromFile'
import { ChartManager } from './manager'
import { isChartType, type ChartLayerOperation, type ChartSpec, type ChartType } from './types'

export interface ChartSnapshotV1 {
  version: 1
  charts: Array<{ spec: ChartSpec; cellAnchor?: FileChartAnchor }>
}

export interface ChartUndoRecord {
  label: string
  before: ChartSnapshotV1
  after: ChartSnapshotV1
}

export interface ChartUndoSink {
  push(record: ChartUndoRecord): void
}

function copy<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T
}

export function isChartSnapshotV1(value: unknown): value is ChartSnapshotV1 {
  if (!value || typeof value !== 'object') return false
  const snapshot = value as Partial<ChartSnapshotV1>
  if (snapshot.version !== 1 || !Array.isArray(snapshot.charts)) return false
  const ids = new Set<string>()
  return snapshot.charts.every((entry) => {
    if (!entry || typeof entry !== 'object' || !entry.spec || typeof entry.spec.id !== 'string' || !entry.spec.id || ids.has(entry.spec.id)) return false
    ids.add(entry.spec.id)
    const range = entry.spec.range
    const anchor = entry.cellAnchor
    const identity = entry.spec.nativeIdentity
    const validIdentity = identity === undefined || (
      /^xl\/charts\/[^/]+\.xml$/.test(identity.part)
      && /^xl\/drawings\/(?!_rels\/)[^/]+\.xml$/.test(identity.drawingPart)
      && Number.isSafeInteger(identity.objectId) && identity.objectId > 0 && identity.objectId <= 0xffff_ffff
    )
    const validAnchor = anchor === undefined || (
      [anchor.FromCol, anchor.FromRow, anchor.ToCol, anchor.ToRow].every(Number.isSafeInteger)
      && anchor.FromCol >= 0 && anchor.FromRow >= 0 && anchor.ToCol > anchor.FromCol && anchor.ToRow > anchor.FromRow
    )
    return isChartType(entry.spec.type) && validIdentity && validAnchor && !!range && typeof range.sheetId === 'string' && range.sheetId.length > 0
      && [range.startRow, range.startColumn, range.endRow, range.endColumn].every(Number.isSafeInteger)
      && range.startRow >= 0 && range.startColumn >= 0
      && range.endRow >= range.startRow && range.endColumn >= range.startColumn
  })
}

export class ChartHandle {
  constructor(private readonly owner: ChartCommandController, readonly id: string) {}

  get value(): ChartSpec | undefined {
    const spec = this.owner.manager.getSpec(this.id)
    return spec ? copy(spec) : undefined
  }

  update(patch: Partial<Omit<ChartSpec, 'id' | 'nativeIdentity'>>): boolean {
    return this.owner.update(this.id, patch)
  }

  remove(): boolean {
    return this.owner.remove(this.id)
  }

  layer(operation: ChartLayerOperation): boolean {
    return this.owner.layer(this.id, operation)
  }
}

/** Host-neutral chart lifecycle facade with complete before/after inverses. */
export class ChartCommandController {
  constructor(readonly manager: ChartManager, private readonly undo?: ChartUndoSink) {}

  createFromSelection(type: ChartType, title?: string): ChartHandle | null {
    const spec = this.record('Create chart', () => this.manager.createFromSelection(type, title))
    return spec ? new ChartHandle(this, spec.id) : null
  }

  add(spec: ChartSpec, cellAnchor?: FileChartAnchor): ChartHandle | null {
    if (!isChartSnapshotV1({ version: 1, charts: [{ spec, cellAnchor }] })) return null
    const added = this.record('Add chart', () => this.manager.add(copy(spec), cellAnchor ? copy(cellAnchor) : undefined))
    return added ? new ChartHandle(this, spec.id) : null
  }

  getById(id: string): ChartHandle | null {
    return this.manager.getSpec(id) ? new ChartHandle(this, id) : null
  }

  list(): ChartHandle[] {
    return this.manager.list().map((spec) => new ChartHandle(this, spec.id))
  }

  update(id: string, patch: Partial<Omit<ChartSpec, 'id' | 'nativeIdentity'>>): boolean {
    const current = this.manager.getSpec(id)
    if (!current) return false
    const untrusted = patch as Partial<ChartSpec>
    const { id: _ignoredId, nativeIdentity: _ignoredIdentity, ...safePatch } = untrusted
    const candidate: ChartSpec = { ...current, ...safePatch, id, nativeIdentity: current.nativeIdentity }
    if (!isChartSnapshotV1({ version: 1, charts: [{ spec: candidate }] })) return false
    this.record('Update chart', () => this.manager.updateSpec(id, copy(safePatch)))
    return true
  }

  remove(id: string): boolean {
    if (!this.manager.getSpec(id)) return false
    this.record('Remove chart', () => this.manager.remove(id))
    return true
  }

  layer(id: string, operation: ChartLayerOperation): boolean {
    if (!['bringForward', 'sendBackward', 'bringToFront', 'sendToBack'].includes(operation)) return false
    if (!this.manager.getSpec(id)) return false
    return this.record(`Layer chart: ${operation}`, () => this.manager.layer(id, operation))
  }

  /** Apply one validated external object mutation without putting a remote
   * operation on the local undo stack. Collaboration never calls restore(). */
  applyExternalCreate(spec: ChartSpec, cellAnchor?: FileChartAnchor): boolean {
    if (!isChartSnapshotV1({ version: 1, charts: [{ spec, cellAnchor }] })) return false
    return this.manager.add(copy(spec), cellAnchor ? copy(cellAnchor) : undefined)
  }

  applyExternalUpdate(id: string, patch: Partial<Omit<ChartSpec, 'id' | 'nativeIdentity'>>): boolean {
    const current = this.manager.getSpec(id)
    if (!current) return false
    const untrusted = patch as Partial<ChartSpec>
    const { id: _ignoredId, nativeIdentity: _ignoredIdentity, ...safePatch } = untrusted
    const candidate: ChartSpec = { ...current, ...safePatch, id, nativeIdentity: current.nativeIdentity }
    if (!isChartSnapshotV1({ version: 1, charts: [{ spec: candidate }] })) return false
    this.manager.updateSpec(id, copy(safePatch))
    return true
  }

  applyExternalRemove(id: string): boolean {
    if (!this.manager.getSpec(id)) return false
    this.manager.remove(id)
    return true
  }

  applyExternalLayer(id: string, operation: ChartLayerOperation): boolean {
    return this.manager.layer(id, operation)
  }

  snapshot(): ChartSnapshotV1 {
    return { version: 1, charts: copy(this.manager.listWithAnchors()) }
  }

  /** Replace all managed charts. Invalid or unmountable snapshots are refused;
   * a best-effort rollback restores the previous complete snapshot. */
  restore(snapshot: ChartSnapshotV1): boolean {
    if (!isChartSnapshotV1(snapshot)) return false
    const before = this.snapshot()
    if (this.applySnapshot(snapshot)) return true
    this.applySnapshot(before)
    return false
  }

  private applySnapshot(snapshot: ChartSnapshotV1): boolean {
    for (const spec of this.manager.list()) this.manager.remove(spec.id)
    for (const entry of snapshot.charts) {
      if (!this.manager.add(copy(entry.spec), entry.cellAnchor ? copy(entry.cellAnchor) : undefined)) return false
    }
    return true
  }

  private record<T>(label: string, mutation: () => T): T {
    const before = this.snapshot()
    const value = mutation()
    const after = this.snapshot()
    if (JSON.stringify(before) !== JSON.stringify(after)) this.undo?.push({ label, before, after })
    return value
  }
}
