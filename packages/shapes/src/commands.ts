import type { FileShapeAnchor } from './fromFile'
import { ShapeManager } from './manager'
import { SHAPE_KINDS, type ShapeSpec, type ShapeKind } from './types'

export interface ShapeSnapshotV1 {
  version: 1
  shapes: ShapeSnapshotEntry[]
}

export interface ShapeSnapshotEntry { spec: ShapeSpec; cellAnchor?: FileShapeAnchor; sheetId: string }

export interface ShapeUndoRecord {
  label: string
  before: ShapeSnapshotV1
  after: ShapeSnapshotV1
}

export interface ShapeUndoSink { push(record: ShapeUndoRecord): void }

const KINDS: ReadonlySet<string> = new Set(SHAPE_KINDS)
const SPEC_FIELDS: ReadonlySet<string> = new Set(['id', 'nativeIdentity', 'kind', 'text', 'fill', 'stroke', 'strokeWidth', 'textColor', 'fontSize'])
const ENTRY_FIELDS: ReadonlySet<string> = new Set(['spec', 'cellAnchor', 'sheetId'])

function copy<T>(value: T): T { return JSON.parse(JSON.stringify(value)) as T }

export function isShapeSnapshot(value: unknown): value is ShapeSnapshotV1 {
  if (!value || typeof value !== 'object') return false
  const snapshot = value as Partial<ShapeSnapshotV1>
  if (Object.keys(value).some((key) => key !== 'version' && key !== 'shapes') || snapshot.version !== 1 || !Array.isArray(snapshot.shapes)) return false
  const ids = new Set<string>()
  return snapshot.shapes.every((entry) => {
    if (!entry || typeof entry !== 'object' || !entry.spec || typeof entry.spec.id !== 'string' || !entry.spec.id || ids.has(entry.spec.id) || !KINDS.has(entry.spec.kind)) return false
    if (Object.keys(entry).some((key) => !ENTRY_FIELDS.has(key))) return false
    if (Object.keys(entry.spec).some((key) => !SPEC_FIELDS.has(key))) return false
    for (const field of ['text', 'fill', 'stroke', 'textColor'] as const) {
      if (entry.spec[field] !== undefined && typeof entry.spec[field] !== 'string') return false
    }
    for (const field of ['strokeWidth', 'fontSize'] as const) {
      const number = entry.spec[field]
      if (number !== undefined && (!Number.isFinite(number) || number < 0)) return false
    }
    ids.add(entry.spec.id)
    if (typeof entry.sheetId !== 'string' || entry.sheetId.length === 0) return false
    const identity = entry.spec.nativeIdentity
    const validIdentity = identity === undefined || (Object.keys(identity).length === 2
      &&
      /^xl\/drawings\/(?!_rels\/)[^/]+\.xml$/.test(identity.drawingPart)
      && Number.isSafeInteger(identity.objectId) && identity.objectId > 0 && identity.objectId <= 0xffff_ffff
    )
    const anchor = entry.cellAnchor
    const validAnchor = anchor === undefined || (Object.keys(anchor).length === 4
      &&
      [anchor.FromCol, anchor.FromRow, anchor.ToCol, anchor.ToRow].every(Number.isSafeInteger)
      && anchor.FromCol >= 0 && anchor.FromRow >= 0 && anchor.ToCol > anchor.FromCol && anchor.ToRow > anchor.FromRow
    )
    return validIdentity && validAnchor
  })
}

export class ShapeHandle {
  constructor(private readonly owner: ShapeCommandController, readonly id: string) {}
  get value(): ShapeSpec | undefined {
    const value = this.owner.manager.getSpec(this.id)
    return value ? copy(value) : undefined
  }
  update(patch: Partial<Omit<ShapeSpec, 'id' | 'kind' | 'nativeIdentity'>>): boolean { return this.owner.update(this.id, patch) }
  remove(): boolean { return this.owner.remove(this.id) }
}

/** Host-neutral shape lifecycle facade with complete snapshot inverses. */
export class ShapeCommandController {
  constructor(readonly manager: ShapeManager, private readonly undo?: ShapeUndoSink) {}

  create(kind: ShapeKind): ShapeHandle | null {
    if (!KINDS.has(kind)) return null
    const spec = this.record('Create shape', () => this.manager.create(kind))
    return spec ? new ShapeHandle(this, spec.id) : null
  }

  add(spec: ShapeSpec, cellAnchor?: FileShapeAnchor, sheetId?: string): ShapeHandle | null {
    const targetSheetId = sheetId ?? this.manager.activeSheetId()
    if (!targetSheetId || !isShapeSnapshot({ version: 1, shapes: [{ spec, cellAnchor, sheetId: targetSheetId }] })) return null
    const added = this.record('Add shape', () => this.manager.add(copy(spec), cellAnchor ? copy(cellAnchor) : undefined, targetSheetId))
    return added ? new ShapeHandle(this, spec.id) : null
  }

  getById(id: string): ShapeHandle | null { return this.manager.getSpec(id) ? new ShapeHandle(this, id) : null }
  list(): ShapeHandle[] { return this.manager.list().map((spec) => new ShapeHandle(this, spec.id)) }

  update(id: string, patch: Partial<Omit<ShapeSpec, 'id' | 'kind' | 'nativeIdentity'>>): boolean {
    const current = this.manager.getSpec(id)
    if (!current) return false
    const untrusted = patch as Partial<ShapeSpec>
    const { id: _id, kind: _kind, nativeIdentity: _identity, ...safePatch } = untrusted
    const candidate: ShapeSpec = { ...current, ...safePatch, id, kind: current.kind, nativeIdentity: current.nativeIdentity }
    const sheetId = this.manager.listWithAnchors().find((entry) => entry.spec.id === id)?.sheetId
    if (!sheetId || !isShapeSnapshot({ version: 1, shapes: [{ spec: candidate, sheetId }] })) return false
    this.record('Update shape', () => this.manager.updateSpec(id, copy(safePatch)))
    return true
  }

  remove(id: string): boolean {
    if (!this.manager.getSpec(id)) return false
    this.record('Remove shape', () => this.manager.remove(id))
    return true
  }

  snapshot(): ShapeSnapshotV1 { return { version: 1, shapes: copy(this.manager.listWithAnchors()) } }

  restore(snapshot: ShapeSnapshotV1): boolean {
    if (!isShapeSnapshot(snapshot)) return false
    const before = this.snapshot()
    if (this.applySnapshot(snapshot)) return true
    this.applySnapshot(before)
    return false
  }

  private applySnapshot(snapshot: ShapeSnapshotV1): boolean {
    for (const spec of this.manager.list()) this.manager.remove(spec.id)
    for (const entry of snapshot.shapes) {
      if (!this.manager.add(copy(entry.spec), entry.cellAnchor ? copy(entry.cellAnchor) : undefined, entry.sheetId)) return false
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
