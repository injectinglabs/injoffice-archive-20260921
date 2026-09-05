import { PivotManager } from './manager'
import { AGG_KINDS, type PivotFieldSort, type PivotMemberFilter, type PivotPageField, type PivotSpec, type PivotValueField } from './types'

export interface PivotSnapshotV1 {
  version: 1
  pivots: PivotSpec[]
}

export interface PivotUndoRecord {
  label: string
  before: PivotSnapshotV1
  after: PivotSnapshotV1
}

export interface PivotUndoSink { push(record: PivotUndoRecord): void }

export type PivotFieldPatch = Partial<Pick<PivotSpec,
  'rows' | 'columns' | 'values' | 'filters' | 'memberFilters' | 'pageFields' | 'sorts'>>

const SPEC_FIELDS = new Set(['id', 'nativeIdentity', 'source', 'rows', 'columns', 'values', 'filters', 'memberFilters', 'pageFields', 'sorts', 'target', 'grandTotals'])
const AGGS = new Set<string>(AGG_KINDS)
const DIRECTIONS = new Set<string>(['ascending', 'descending'])

function copy<T>(value: T): T { return structuredClone(value) }

function stringArray(value: unknown): value is string[] {
  return Array.isArray(value) && value.every((item) => typeof item === 'string')
}

function object(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === 'object' && !Array.isArray(value) && Object.getPrototypeOf(value) === Object.prototype
}

function exactKeys(value: object, required: readonly string[], optional: readonly string[] = []): boolean {
  const keys = Object.keys(value)
  return required.every((key) => keys.includes(key)) && keys.every((key) => required.includes(key) || optional.includes(key))
}

function uniqueFields(values: readonly unknown[]): boolean {
  const fields = values.map((value) => object(value) ? value.field : undefined)
  return fields.every((field) => typeof field === 'string') && new Set(fields).size === values.length
}

function validRange(value: PivotSpec['source']): boolean {
  return !!value && exactKeys(value, ['sheetId', 'startRow', 'startColumn', 'endRow', 'endColumn'])
    && typeof value.sheetId === 'string' && value.sheetId.length > 0
    && [value.startRow, value.startColumn, value.endRow, value.endColumn].every(Number.isSafeInteger)
    && value.startRow >= 0 && value.startColumn >= 0
    && value.endRow >= value.startRow && value.endColumn >= value.startColumn
}

function validSpec(spec: PivotSpec): boolean {
  if (!spec || typeof spec !== 'object' || Object.keys(spec).some((key) => !SPEC_FIELDS.has(key))) return false
  if (typeof spec.id !== 'string' || spec.id.length === 0 || !validRange(spec.source)) return false
  if (!spec.target || !exactKeys(spec.target, ['sheetId', 'startRow', 'startColumn'])
    || typeof spec.target.sheetId !== 'string' || spec.target.sheetId.length === 0
    || !Number.isSafeInteger(spec.target.startRow) || spec.target.startRow < 0
    || !Number.isSafeInteger(spec.target.startColumn) || spec.target.startColumn < 0) return false
  if (!stringArray(spec.rows) || !stringArray(spec.columns) || spec.columns.length > 1
    || new Set(spec.rows).size !== spec.rows.length || new Set(spec.columns).size !== spec.columns.length
    || new Set([...spec.rows, ...spec.columns]).size !== spec.rows.length + spec.columns.length) return false
  if (!Array.isArray(spec.values) || spec.values.length === 0 || !spec.values.every((value) => object(value) && exactKeys(value, ['field', 'agg'])
    && typeof value.field === 'string' && value.field.length > 0 && typeof value.agg === 'string' && AGGS.has(value.agg))) return false
  if (spec.grandTotals !== undefined && typeof spec.grandTotals !== 'boolean') return false
  if (spec.filters !== undefined && (!spec.filters || typeof spec.filters !== 'object' || Array.isArray(spec.filters)
    || !Object.entries(spec.filters).every(([field, values]) => field.length > 0 && stringArray(values)))) return false
  if (spec.memberFilters !== undefined && (!Array.isArray(spec.memberFilters) || !uniqueFields(spec.memberFilters)
    || !spec.memberFilters.every((filter) => object(filter) && exactKeys(filter, ['field', 'mode', 'values']) && typeof filter.field === 'string' && filter.field.length > 0
      && (filter.mode === 'include' || filter.mode === 'exclude') && stringArray(filter.values)))) return false
  if (spec.pageFields !== undefined && (!Array.isArray(spec.pageFields) || !uniqueFields(spec.pageFields)
    || !spec.pageFields.every((page) => object(page) && exactKeys(page, ['field'], ['selectedItem']) && typeof page.field === 'string' && page.field.length > 0
      && (page.selectedItem === undefined || typeof page.selectedItem === 'string')))) return false
  if (spec.sorts !== undefined && (!Array.isArray(spec.sorts) || !uniqueFields(spec.sorts)
    || !spec.sorts.every((sort) => object(sort) && exactKeys(sort, ['field', 'direction']) && typeof sort.field === 'string' && sort.field.length > 0
      && typeof sort.direction === 'string' && DIRECTIONS.has(sort.direction)))) return false
  const axes = new Set([...spec.rows, ...spec.columns])
  if ((spec.pageFields ?? []).some(({ field, selectedItem }) => axes.has(field)
    || (selectedItem !== undefined && (!!spec.filters?.[field] || (spec.memberFilters ?? []).some((filter) => filter.field === field))))) return false
  if ((spec.sorts ?? []).some(({ field }) => !axes.has(field))) return false
  const identity = spec.nativeIdentity
  return identity === undefined || (object(identity) && exactKeys(identity, ['part']) && typeof identity.part === 'string'
    && /^xl\/pivotTables\/(?!_rels\/)[^/]+\.xml$/.test(identity.part))
}

export function isPivotSnapshot(value: unknown): value is PivotSnapshotV1 {
  if (!value || typeof value !== 'object') return false
  const snapshot = value as Partial<PivotSnapshotV1>
  if (!exactKeys(value, ['version', 'pivots']) || snapshot.version !== 1 || !Array.isArray(snapshot.pivots)) return false
  const ids = new Set<string>()
  return snapshot.pivots.every((spec) => {
    if (!validSpec(spec) || ids.has(spec.id)) return false
    ids.add(spec.id)
    return true
  })
}

export class PivotHandle {
  constructor(private readonly owner: PivotCommandController, readonly id: string) {}
  get value(): PivotSpec | undefined {
    const value = this.owner.manager.getSpec(this.id)
    return value ? copy(value) : undefined
  }
  update(patch: Partial<Omit<PivotSpec, 'id' | 'nativeIdentity'>>): boolean { return this.owner.update(this.id, patch) }
  setRowFields(rows: string[]): boolean { return this.owner.setRowFields(this.id, rows) }
  setColumnFields(columns: string[]): boolean { return this.owner.setColumnFields(this.id, columns) }
  setValueFields(values: PivotValueField[]): boolean { return this.owner.setValueFields(this.id, values) }
  setFilters(filters?: Record<string, string[]>): boolean { return this.owner.setFilters(this.id, filters) }
  setMemberFilters(memberFilters?: PivotMemberFilter[]): boolean { return this.owner.setMemberFilters(this.id, memberFilters) }
  setPageFields(pageFields?: PivotPageField[]): boolean { return this.owner.setPageFields(this.id, pageFields) }
  setSorts(sorts?: PivotFieldSort[]): boolean { return this.owner.setSorts(this.id, sorts) }
  remove(): boolean { return this.owner.remove(this.id) }
}

/** Host-neutral pivot lifecycle and field facade with complete snapshot inverses. */
export class PivotCommandController {
  constructor(readonly manager: PivotManager, private readonly undo?: PivotUndoSink) {}

  createFromSelection(): PivotHandle | null {
    const spec = this.record('Create pivot', () => this.manager.createFromSelection())
    return spec ? new PivotHandle(this, spec.id) : null
  }

  add(spec: PivotSpec): PivotHandle | null {
    if (!this.validForHost(spec)) return null
    const added = this.record('Add pivot', () => this.manager.add(copy(spec)))
    return added ? new PivotHandle(this, spec.id) : null
  }

  getById(id: string): PivotHandle | null { return this.manager.getSpec(id) ? new PivotHandle(this, id) : null }
  list(): PivotHandle[] { return this.manager.list().map(({ id }) => new PivotHandle(this, id)) }

  update(id: string, patch: Partial<Omit<PivotSpec, 'id' | 'nativeIdentity'>>): boolean {
    return this.updateWithLabel(id, patch, 'Update pivot')
  }

  setRowFields(id: string, rows: string[]): boolean { return this.updateWithLabel(id, { rows }, 'Set pivot row fields') }
  setColumnFields(id: string, columns: string[]): boolean { return this.updateWithLabel(id, { columns }, 'Set pivot column fields') }
  setValueFields(id: string, values: PivotValueField[]): boolean { return this.updateWithLabel(id, { values }, 'Set pivot value fields') }
  setFilters(id: string, filters?: Record<string, string[]>): boolean { return this.updateWithLabel(id, { filters }, 'Set pivot filters') }
  setMemberFilters(id: string, memberFilters?: PivotMemberFilter[]): boolean { return this.updateWithLabel(id, { memberFilters }, 'Set pivot member filters') }
  setPageFields(id: string, pageFields?: PivotPageField[]): boolean { return this.updateWithLabel(id, { pageFields }, 'Set pivot page fields') }
  setSorts(id: string, sorts?: PivotFieldSort[]): boolean { return this.updateWithLabel(id, { sorts }, 'Set pivot sorts') }

  remove(id: string): boolean {
    if (!this.manager.getSpec(id)) return false
    this.record('Remove pivot', () => this.manager.remove(id))
    return true
  }

  snapshot(): PivotSnapshotV1 { return { version: 1, pivots: copy(this.manager.list()) } }

  restore(snapshot: PivotSnapshotV1): boolean {
    if (!isPivotSnapshot(snapshot) || !snapshot.pivots.every((spec) => this.validForHost(spec))) return false
    const before = this.snapshot()
    if (this.applySnapshot(snapshot)) return true
    this.applySnapshot(before)
    return false
  }

  private validForHost(spec: PivotSpec): boolean {
    if (!validSpec(spec) || !this.manager.canMount(spec)) return false
    const sourceFields = new Set(this.manager.sourceFieldsAt(spec.source))
    const configured = [
      ...spec.rows, ...spec.columns, ...spec.values.map(({ field }) => field),
      ...Object.keys(spec.filters ?? {}), ...(spec.memberFilters ?? []).map(({ field }) => field),
      ...(spec.pageFields ?? []).map(({ field }) => field), ...(spec.sorts ?? []).map(({ field }) => field),
    ]
    return sourceFields.size > 0 && configured.every((field) => sourceFields.has(field))
  }

  private updateWithLabel(id: string, patch: Partial<Omit<PivotSpec, 'id' | 'nativeIdentity'>>, label: string): boolean {
    const current = this.manager.getSpec(id)
    if (!current || !patch || typeof patch !== 'object') return false
    const untrusted = patch as Partial<PivotSpec>
    const { id: _id, nativeIdentity: _identity, ...safePatch } = untrusted
    const candidate: PivotSpec = { ...current, ...safePatch, id, nativeIdentity: current.nativeIdentity }
    if (!this.validForHost(candidate)) return false
    this.record(label, () => this.manager.updateSpec(id, copy(safePatch)))
    return true
  }

  private applySnapshot(snapshot: PivotSnapshotV1): boolean {
    for (const spec of this.manager.list()) this.manager.remove(spec.id)
    for (const spec of snapshot.pivots) if (!this.manager.add(copy(spec))) return false
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
