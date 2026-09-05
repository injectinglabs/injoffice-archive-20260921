import {
  ConnectorManager,
  isConnectorManagerSnapshot,
  isConnectorRangeSnapshot,
  type ConnectorManagerSnapshotV1,
  type ConnectorRangeSnapshot,
} from './manager'
import type { ConnectorRefreshOptions, ConnectorRefreshResult, ConnectorSpec } from './types'

export interface ConnectorCommandSnapshotV1 {
  version: 1
  manager: ConnectorManagerSnapshotV1
  ranges: ConnectorRangeSnapshot[]
}

export interface ConnectorUndoRecord {
  label: string
  before: ConnectorCommandSnapshotV1
  after: ConnectorCommandSnapshotV1
}

export interface ConnectorUndoSink { push(record: ConnectorUndoRecord): void }

export interface ConnectorCommandControllerOptions {
  /** Optional extra host policy. The built-in credential-free URL check still applies. */
  canPersist?: (spec: Readonly<ConnectorSpec>) => boolean
}

export interface ConnectorRestoreOptions {
  /** Required when the manager is remotely authoritative, preventing ordinary
   * local undo/redo from bypassing the ordered connector operation stream. */
  source?: 'collaboration'
}

function clone<T>(value: T): T { return structuredClone(value) }

function onlyKeys(value: object, allowed: readonly string[]): boolean {
  const keys = new Set(allowed)
  return Object.keys(value).every((key) => keys.has(key))
}

/** Command and undo snapshots intentionally reject URLs that can carry inline
 * credentials, signed query data, or fragments. Hosts should use an opaque,
 * credential-free gateway URL and keep authorization in SourceFetcher. */
export function isCredentialFreeConnectorSpec(spec: unknown): spec is ConnectorSpec {
  if (!spec || typeof spec !== 'object') return false
  const value = spec as Partial<ConnectorSpec>
  if (!value.source || typeof value.source !== 'object' || !value.target || typeof value.target !== 'object'
    || value.source.kind !== 'http' || typeof value.source.url !== 'string' || !value.source.url.trim()) return false
  try {
    const url = new URL(value.source.url, 'https://injoffice.invalid')
    return onlyKeys(value, ['id', 'name', 'source', 'target', 'refresh', 'schedule', 'cache', 'schema'])
      && onlyKeys(value.source, ['kind', 'url', 'format', 'path'])
      && onlyKeys(value.target, ['sheetId', 'startRow', 'startColumn'])
      && (!value.schedule || onlyKeys(value.schedule, ['intervalMs']))
      && (!value.cache || onlyKeys(value.cache, ['mode', 'ttlMs']))
      && (!value.schema || onlyKeys(value.schema, ['columns', 'allowAdditionalColumns'])
        && Array.isArray(value.schema.columns)
        && value.schema.columns.every((column) => onlyKeys(column, ['index', 'type', 'nullable'])))
      && (url.protocol === 'http:' || url.protocol === 'https:')
      && !url.username && !url.password && !url.search && !url.hash
  } catch {
    return false
  }
}

function same(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}

export class ConnectorHandle {
  constructor(private readonly owner: ConnectorCommandController, readonly id: string) {}
  get value(): ConnectorSpec | undefined { return this.owner.manager.get(this.id) }
  update(patch: Partial<Omit<ConnectorSpec, 'id'>>): boolean { return this.owner.update(this.id, patch) }
  refresh(options?: ConnectorRefreshOptions): Promise<ConnectorRefreshResult> { return this.owner.refresh(this.id, options) }
  remove(): boolean { return this.owner.remove(this.id) }
}

/** Host-neutral lifecycle commands with complete model and bounded cell-range
 * inverses. Refresh fetches only once; redo replays captured cells, never an
 * external request. */
export class ConnectorCommandController {
  private readonly canPersist: (spec: Readonly<ConnectorSpec>) => boolean

  constructor(
    readonly manager: ConnectorManager,
    private readonly undo?: ConnectorUndoSink,
    options: ConnectorCommandControllerOptions = {},
  ) {
    this.canPersist = options.canPersist ?? (() => true)
  }

  getById(id: string): ConnectorHandle | null {
    return this.manager.get(id) ? new ConnectorHandle(this, id) : null
  }

  list(): ConnectorHandle[] {
    return this.manager.list().map((spec) => new ConnectorHandle(this, spec.id))
  }

  add(spec: ConnectorSpec): ConnectorHandle | null {
    if (!this.persistable(spec)) return null
    const before = this.snapshot()
    try {
      if (!this.manager.add(clone(spec))) return null
    } catch {
      return null
    }
    this.push('Add data connection', before, this.snapshot())
    return new ConnectorHandle(this, spec.id)
  }

  async createAtSelection(
    input: Omit<ConnectorSpec, 'id' | 'target'>,
    refreshOptions: ConnectorRefreshOptions = {},
  ): Promise<ConnectorHandle | null> {
    const beforeState = this.manager.snapshotState()
    let created: ConnectorSpec | null
    try {
      created = this.manager.addAtSelection(clone(input))
    } catch {
      return null
    }
    if (!created) return null
    if (!this.persistable(created)) {
      this.manager.restoreState(beforeState)
      return null
    }
    try {
      const transaction = await this.manager.refreshTransaction(created.id, refreshOptions)
      if (transaction.result !== 'applied' || !transaction.afterState) {
        this.manager.restoreState(beforeState)
        return null
      }
      const before: ConnectorCommandSnapshotV1 = {
        version: 1,
        manager: beforeState,
        ranges: transaction.beforeRange ? [transaction.beforeRange] : [],
      }
      const after: ConnectorCommandSnapshotV1 = {
        version: 1,
        manager: transaction.afterState,
        ranges: transaction.afterRange ? [transaction.afterRange] : [],
      }
      this.push('Create data connection', before, after)
      return new ConnectorHandle(this, created.id)
    } catch (error) {
      this.manager.restoreState(beforeState)
      throw error
    }
  }

  update(id: string, patch: Partial<Omit<ConnectorSpec, 'id'>>): boolean {
    const current = this.manager.get(id)
    if (!current) return false
    const candidate = { ...current, ...clone(patch), id }
    if (!this.persistable(candidate)) return false
    if (same(current, candidate)) return true
    const beforeState = this.manager.snapshotState()
    const beforeRange = this.manager.captureExtent(id)
    const beforeEntry = beforeState.connectors.find(({ spec }) => spec.id === id)
    if (beforeEntry && beforeEntry.lastRows > 0 && beforeEntry.lastColumns > 0 && !beforeRange) return false
    try {
      if (!this.manager.update(id, patch)) return false
      const afterRange = beforeRange
        ? this.manager.captureRangeSnapshot(beforeRange, beforeRange.rowCount, beforeRange.columnCount)
        : null
      if (beforeRange && !afterRange) throw new Error('connector update could not capture the resulting worksheet range')
      this.push('Update data connection',
        { version: 1, manager: beforeState, ranges: beforeRange ? [beforeRange] : [] },
        { version: 1, manager: this.manager.snapshotState(), ranges: afterRange ? [afterRange] : [] })
      return true
    } catch (error) {
      if (beforeRange) this.manager.applyRangeSnapshot(beforeRange)
      this.manager.restoreState(beforeState)
      if (error instanceof TypeError) return false
      throw error
    }
  }

  remove(id: string): boolean {
    if (!this.manager.get(id)) return false
    const beforeState = this.manager.snapshotState()
    const beforeRange = this.manager.captureExtent(id)
    const beforeEntry = beforeState.connectors.find(({ spec }) => spec.id === id)
    if (beforeEntry && beforeEntry.lastRows > 0 && beforeEntry.lastColumns > 0 && !beforeRange) return false
    try {
      if (!this.manager.remove(id)) return false
      const afterRange = beforeRange
        ? this.manager.captureRangeSnapshot(beforeRange, beforeRange.rowCount, beforeRange.columnCount)
        : null
      if (beforeRange && !afterRange) throw new Error('connector removal could not capture the cleared worksheet range')
      this.push('Remove data connection',
        { version: 1, manager: beforeState, ranges: beforeRange ? [beforeRange] : [] },
        { version: 1, manager: this.manager.snapshotState(), ranges: afterRange ? [afterRange] : [] })
      return true
    } catch (error) {
      if (beforeRange) this.manager.applyRangeSnapshot(beforeRange)
      this.manager.restoreState(beforeState)
      throw error
    }
  }

  async refresh(id: string, options: ConnectorRefreshOptions = {}): Promise<ConnectorRefreshResult> {
    const transaction = await this.manager.refreshTransaction(id, options)
    if (transaction.result !== 'applied' || !transaction.afterState) return transaction.result
    this.push('Refresh data connection',
      { version: 1, manager: transaction.beforeState, ranges: transaction.beforeRange ? [transaction.beforeRange] : [] },
      { version: 1, manager: transaction.afterState, ranges: transaction.afterRange ? [transaction.afterRange] : [] })
    return transaction.result
  }

  snapshot(): ConnectorCommandSnapshotV1 {
    return { version: 1, manager: this.manager.snapshotState(), ranges: [] }
  }

  restore(snapshot: ConnectorCommandSnapshotV1, options: ConnectorRestoreOptions = {}): boolean {
    if (this.manager.executionAuthority === 'server' && options.source !== 'collaboration') return false
    if (!this.validSnapshot(snapshot)) return false
    const liveState = this.manager.snapshotState()
    const liveRanges: ConnectorRangeSnapshot[] = []
    for (const range of snapshot.ranges) {
      const live = this.manager.captureRangeSnapshot(range, range.rowCount, range.columnCount)
      if (!live) return false
      liveRanges.push(live)
    }
    let attempted = -1
    try {
      for (let index = 0; index < snapshot.ranges.length; index++) {
        attempted = index
        const range = snapshot.ranges[index]
        if (!this.manager.applyRangeSnapshot(range, options)) throw new Error('connector worksheet snapshot could not be restored')
      }
      if (!this.manager.restoreState(snapshot.manager, options)) throw new Error('connector model snapshot could not be restored')
      return true
    } catch {
      // Restore the currently attempted range too: a host setValues adapter
      // may write a prefix and then throw rather than failing atomically.
      for (let index = attempted; index >= 0; index--) this.manager.applyRangeSnapshot(liveRanges[index], options)
      this.manager.restoreState(liveState, options)
      return false
    }
  }

  private persistable(spec: ConnectorSpec): boolean {
    if (!isCredentialFreeConnectorSpec(spec)) return false
    try {
      return this.canPersist(clone(spec))
    } catch {
      return false
    }
  }

  private validSnapshot(snapshot: ConnectorCommandSnapshotV1): boolean {
    return !!snapshot && snapshot.version === 1 && isConnectorManagerSnapshot(snapshot.manager)
      && snapshot.manager.connectors.every(({ spec }) => this.persistable(spec))
      && Array.isArray(snapshot.ranges) && snapshot.ranges.every(isConnectorRangeSnapshot)
  }

  private push(label: string, before: ConnectorCommandSnapshotV1, after: ConnectorCommandSnapshotV1): void {
    if (!same(before, after)) this.undo?.push({ label, before: clone(before), after: clone(after) })
  }
}
