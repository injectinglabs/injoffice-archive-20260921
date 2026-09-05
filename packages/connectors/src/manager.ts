import type { ICellData } from '@univerjs/core'
import type { FUniver } from '@univerjs/core/lib/facade'
import type {} from '@univerjs/sheets/lib/facade'
import type {} from '@univerjs/sheets-ui/lib/facade'
import type {} from '@univerjs/ui/lib/facade'
import { RangePreprocessError, RangePreprocessPipeline, type RangeValue } from './preprocess'
import type {
  ConnectorAuthorizer,
  ConnectorRefreshOptions,
  ConnectorRefreshResult,
  ConnectorScheduler,
  ConnectorSpec,
  ConnectorStatus,
  SourceFetcher,
} from './types'

let seq = 0
function newConnectorId(): string {
  return `conn-${Date.now().toString(36)}-${(++seq).toString(36)}`
}

interface MountedConnector {
  spec: ConnectorSpec
  status: ConnectorStatus
  lastRows: number
  lastCols: number
  controller?: AbortController
  timer?: unknown
}

export interface ConnectorMountedSnapshot {
  spec: ConnectorSpec
  status: Omit<ConnectorStatus, 'lastError' | 'refreshing'>
  lastRows: number
  lastColumns: number
}

export interface ConnectorManagerSnapshotV1 {
  version: 1
  connectors: ConnectorMountedSnapshot[]
}

export interface ConnectorRangeSnapshot {
  sheetId: string
  startRow: number
  startColumn: number
  rowCount: number
  columnCount: number
  values: ICellData[][]
}

export interface ConnectorRefreshTransaction {
  result: ConnectorRefreshResult
  beforeState: ConnectorManagerSnapshotV1
  afterState?: ConnectorManagerSnapshotV1
  beforeRange?: ConnectorRangeSnapshot
  afterRange?: ConnectorRangeSnapshot
}

export interface ConnectorAuthoritativeApplyOptions { source?: 'collaboration' }

interface CacheEntry {
  expiresAt: number
  grid: RangeValue[][]
}

export interface ConnectorManagerOptions {
  authorize?: ConnectorAuthorizer
  preprocessing?: RangePreprocessPipeline
  scheduler?: ConnectorScheduler
  now?: () => number
  /** In server mode local lifecycle mutation, fetch, on-open refresh, and
   * interval timers are disabled. An ordered collaboration session applies
   * authoritative snapshots through restoreState instead. */
  executionAuthority?: 'client' | 'server'
}

const browserScheduler: ConnectorScheduler = {
  setInterval: (callback, intervalMs) => globalThis.setInterval(callback, intervalMs),
  clearInterval: (handle) => globalThis.clearInterval(handle as ReturnType<typeof setInterval>),
}

function plainObject(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === 'object' && !Array.isArray(value) && Object.getPrototypeOf(value) === Object.prototype
}

function onlyKeys(value: object, allowed: readonly string[]): boolean {
  return Object.keys(value).every((key) => allowed.includes(key))
}

function validateSpec(spec: ConnectorSpec): void {
  if (!plainObject(spec) || !onlyKeys(spec, ['id', 'name', 'source', 'target', 'refresh', 'schedule', 'cache', 'schema'])
    || typeof spec.id !== 'string' || !spec.id || typeof spec.name !== 'string' || !spec.name
    || !plainObject(spec.source) || !onlyKeys(spec.source, ['kind', 'url', 'format', 'path'])
    || spec.source.kind !== 'http' || typeof spec.source.url !== 'string' || !spec.source.url
    || !['json', 'csv'].includes(spec.source.format)
    || (spec.source.path !== undefined && typeof spec.source.path !== 'string')
    || !plainObject(spec.target) || !onlyKeys(spec.target, ['sheetId', 'startRow', 'startColumn'])
    || typeof spec.target.sheetId !== 'string' || !spec.target.sheetId) throw new TypeError('connector id, name, source, and target sheet are required')
  if (!['manual', 'onOpen', 'interval'].includes(spec.refresh)) throw new TypeError('unsupported refresh policy')
  if (![spec.target.startRow, spec.target.startColumn].every((value) => Number.isSafeInteger(value) && value >= 0)) throw new TypeError('connector target coordinates must be non-negative integers')
  if (spec.schedule && (!plainObject(spec.schedule) || !onlyKeys(spec.schedule, ['intervalMs']))) throw new TypeError('connector schedule is invalid')
  if (spec.refresh === 'interval' && (!spec.schedule || !Number.isSafeInteger(spec.schedule.intervalMs) || spec.schedule.intervalMs < 1_000)) throw new TypeError('interval refresh requires a schedule of at least 1000ms')
  if (spec.refresh !== 'interval' && spec.schedule) throw new TypeError('schedule is only valid for interval refresh')
  if (spec.cache && (!plainObject(spec.cache) || !onlyKeys(spec.cache, ['mode', 'ttlMs']) || spec.cache.mode !== 'none' && spec.cache.mode !== 'memory' || !Number.isSafeInteger(spec.cache.ttlMs) || spec.cache.ttlMs < 0)) throw new TypeError('cache policy is invalid')
  if (spec.schema && (!plainObject(spec.schema) || !onlyKeys(spec.schema, ['columns', 'allowAdditionalColumns']) || !Array.isArray(spec.schema.columns)
    || spec.schema.allowAdditionalColumns !== undefined && typeof spec.schema.allowAdditionalColumns !== 'boolean')) throw new TypeError('connector schema is invalid')
  const indexes = new Set<number>()
  for (const column of spec.schema?.columns ?? []) {
    if (!plainObject(column) || !onlyKeys(column, ['index', 'type', 'nullable']) || !Number.isSafeInteger(column.index) || column.index < 0 || indexes.has(column.index)
      || !['string', 'number', 'boolean', 'any'].includes(column.type) || column.nullable !== undefined && typeof column.nullable !== 'boolean') throw new TypeError('schema columns must have unique indexes and supported types')
    indexes.add(column.index)
  }
}

function normalizeGrid(value: unknown[][]): RangeValue[][] {
  if (!Array.isArray(value)) throw new TypeError('connector data must be a two-dimensional array')
  const width = value.length ? Math.max(...value.map((row) => Array.isArray(row) ? row.length : -1)) : 0
  if (width < 0) throw new TypeError('connector data rows must be arrays')
  return value.map((row) => {
    if (!Array.isArray(row)) throw new TypeError('connector data rows must be arrays')
    return Array.from({ length: width }, (_, index) => {
      const cell = row[index] ?? null
      if (cell !== null && typeof cell !== 'string' && typeof cell !== 'boolean' && (typeof cell !== 'number' || !Number.isFinite(cell))) throw new TypeError(`unsupported connector cell at column ${index}`)
      return cell
    })
  })
}

function validateSchema(grid: RangeValue[][], spec: ConnectorSpec): void {
  const schema = spec.schema
  if (!schema) return
  const width = grid[0]?.length ?? 0
  const declaredWidth = schema.columns.reduce((maximum, column) => Math.max(maximum, column.index + 1), 0)
  if (!schema.allowAdditionalColumns && width > declaredWidth) throw new TypeError(`connector returned ${width} columns but schema declares ${declaredWidth}`)
  for (const column of schema.columns) {
    for (let row = 0; row < grid.length; row++) {
      const value = grid[row][column.index] ?? null
      if (value === null) {
        if (!column.nullable) throw new TypeError(`row ${row} column ${column.index} is null`)
      } else if (column.type !== 'any' && typeof value !== column.type) {
        throw new TypeError(`row ${row} column ${column.index} must be ${column.type}`)
      }
    }
  }
}

function cloneGrid(grid: RangeValue[][]): RangeValue[][] {
  return grid.map((row) => [...row])
}

function clone<T>(value: T): T {
  return structuredClone(value)
}

function validDimension(value: unknown): value is number {
  return Number.isSafeInteger(value) && (value as number) >= 0
}

export function isConnectorManagerSnapshot(snapshot: unknown): snapshot is ConnectorManagerSnapshotV1 {
  if (!snapshot || typeof snapshot !== 'object') return false
  const value = snapshot as Partial<ConnectorManagerSnapshotV1>
  if (Object.keys(value).some((key) => key !== 'version' && key !== 'connectors') || value.version !== 1 || !Array.isArray(value.connectors)) return false
  const ids = new Set<string>()
  try {
    return value.connectors.every((entry) => {
      if (!entry || typeof entry !== 'object' || !validDimension(entry.lastRows) || !validDimension(entry.lastColumns)) return false
      if (Object.keys(entry).some((key) => !['spec', 'status', 'lastRows', 'lastColumns'].includes(key))) return false
      validateSpec(entry.spec)
      if (ids.has(entry.spec.id)) return false
      ids.add(entry.spec.id)
      if (!entry.status || typeof entry.status !== 'object') return false
      const statusKeys = new Set(['lastRefreshTs', 'rowCount', 'columnCount', 'nextRefreshTs', 'cacheHit', 'lastErrorCode'])
      if (Object.keys(entry.status).some((key) => !statusKeys.has(key))) return false
      for (const value of [entry.status.lastRefreshTs, entry.status.rowCount, entry.status.columnCount, entry.status.nextRefreshTs]) {
        if (value !== undefined && (!Number.isFinite(value) || value < 0)) return false
      }
      if (entry.status.cacheHit !== undefined && typeof entry.status.cacheHit !== 'boolean') return false
      return entry.status.lastErrorCode === undefined
        || ['ABORTED', 'AUTHORIZATION_DENIED', 'FETCH_FAILED', 'INVALID_DATA', 'PREPROCESS_FAILED', 'TARGET_MISSING'].includes(entry.status.lastErrorCode)
    })
  } catch {
    return false
  }
}

export function isConnectorRangeSnapshot(snapshot: unknown): snapshot is ConnectorRangeSnapshot {
  if (!snapshot || typeof snapshot !== 'object') return false
  const value = snapshot as Partial<ConnectorRangeSnapshot>
  return Object.keys(value).every((key) => ['sheetId', 'startRow', 'startColumn', 'rowCount', 'columnCount', 'values'].includes(key))
    && typeof value.sheetId === 'string' && value.sheetId.length > 0
    && validDimension(value.startRow) && validDimension(value.startColumn)
    && validDimension(value.rowCount) && validDimension(value.columnCount)
    && value.rowCount > 0 && value.columnCount > 0 && Array.isArray(value.values)
    && value.values.length === value.rowCount
    && value.values.every((row) => Array.isArray(row) && row.length === value.columnCount)
}

/** Bound-range lifecycle with host-owned authorization, transport, and scheduling. */
export class ConnectorManager {
  private readonly connectors = new Map<string, MountedConnector>()
  private readonly cache = new Map<string, CacheEntry>()
  private readonly changeListeners = new Set<() => void>()
  private readonly authorize?: ConnectorAuthorizer
  private readonly preprocessing?: RangePreprocessPipeline
  private readonly scheduler: ConnectorScheduler
  private readonly now: () => number
  readonly executionAuthority: 'client' | 'server'

  constructor(
    private readonly api: FUniver,
    private readonly fetchSource: SourceFetcher,
    options: ConnectorManagerOptions = {},
  ) {
    this.authorize = options.authorize
    this.preprocessing = options.preprocessing
    this.scheduler = options.scheduler ?? browserScheduler
    this.now = options.now ?? Date.now
    this.executionAuthority = options.executionAuthority ?? 'client'
  }

  onChange(listener: () => void): () => void {
    this.changeListeners.add(listener)
    return () => this.changeListeners.delete(listener)
  }

  private emitChange(): void {
    this.changeListeners.forEach((listener) => listener())
  }

  async createAtSelection(spec: Omit<ConnectorSpec, 'id' | 'target'>): Promise<ConnectorSpec | null> {
    const full = this.addAtSelection(spec)
    if (!full) return null
    await this.refresh(full.id)
    return clone(full)
  }

  addAtSelection(spec: Omit<ConnectorSpec, 'id' | 'target'>): ConnectorSpec | null {
    const workbook = this.api.getActiveWorkbook()
    const sheet = workbook?.getActiveSheet()
    const range = sheet?.getActiveRange() ?? workbook?.getActiveRange()
    if (!workbook || !sheet) return null
    const selected = range?.getRange()
    const full: ConnectorSpec = {
      ...spec,
      id: newConnectorId(),
      target: { sheetId: sheet.getSheetId(), startRow: selected?.startRow ?? 0, startColumn: selected?.startColumn ?? 0 },
    }
    if (!this.add(full)) return null
    return clone(full)
  }

  add(spec: ConnectorSpec): boolean {
    if (this.executionAuthority === 'server') return false
    validateSpec(spec)
    if (this.connectors.has(spec.id)) return false
    const mounted: MountedConnector = { spec: structuredClone(spec), status: {}, lastRows: 0, lastCols: 0 }
    this.connectors.set(spec.id, mounted)
    this.startSchedule(mounted)
    this.emitChange()
    return true
  }

  update(id: string, patch: Partial<Omit<ConnectorSpec, 'id'>>): boolean {
    if (this.executionAuthority === 'server') return false
    const connector = this.connectors.get(id)
    if (!connector) return false
    const candidate = { ...connector.spec, ...clone(patch), id }
    validateSpec(candidate)
    const moved = candidate.target.sheetId !== connector.spec.target.sheetId
      || candidate.target.startRow !== connector.spec.target.startRow
      || candidate.target.startColumn !== connector.spec.target.startColumn
    if (moved) this.clearExtent(connector)
    this.stop(connector)
    connector.spec = candidate
    if (moved) {
      connector.lastRows = 0
      connector.lastCols = 0
      connector.status = {}
    }
    this.startSchedule(connector)
    this.cache.delete(id)
    this.emitChange()
    return true
  }

  remove(id: string): boolean {
    if (this.executionAuthority === 'server') return false
    const connector = this.connectors.get(id)
    if (!connector) return false
    this.stop(connector)
    this.clearExtent(connector)
    this.cache.delete(id)
    this.connectors.delete(id)
    this.emitChange()
    return true
  }

  dispose(): void {
    for (const connector of this.connectors.values()) this.stop(connector)
    this.cache.clear()
    this.changeListeners.clear()
  }

  cancel(id: string): boolean {
    const connector = this.connectors.get(id)
    if (!connector?.controller) return false
    connector.controller.abort('connector refresh canceled')
    return true
  }

  list(): ConnectorSpec[] {
    return [...this.connectors.values()].map((connector) => structuredClone(connector.spec))
  }

  get(id: string): ConnectorSpec | undefined {
    const connector = this.connectors.get(id)
    return connector ? clone(connector.spec) : undefined
  }

  snapshotState(): ConnectorManagerSnapshotV1 {
    return {
      version: 1,
      connectors: [...this.connectors.values()].map((connector) => {
        const { lastError: _lastError, refreshing: _refreshing, ...status } = connector.status
        return {
          spec: clone(connector.spec),
          status: clone(status),
          lastRows: connector.lastRows,
          lastColumns: connector.lastCols,
        }
      }),
    }
  }

  restoreState(snapshot: ConnectorManagerSnapshotV1, options: ConnectorAuthoritativeApplyOptions = {}): boolean {
    if (this.executionAuthority === 'server' && options.source !== 'collaboration') return false
    if (!isConnectorManagerSnapshot(snapshot)) return false
    for (const connector of this.connectors.values()) this.stop(connector)
    this.connectors.clear()
    this.cache.clear()
    for (const entry of snapshot.connectors) {
      const mounted: MountedConnector = {
        spec: clone(entry.spec),
        status: clone(entry.status),
        lastRows: entry.lastRows,
        lastCols: entry.lastColumns,
      }
      this.connectors.set(mounted.spec.id, mounted)
      this.startSchedule(mounted)
    }
    this.emitChange()
    return true
  }

  captureExtent(id: string): ConnectorRangeSnapshot | null {
    const connector = this.connectors.get(id)
    if (!connector || !connector.lastRows || !connector.lastCols) return null
    return this.captureRange(connector.spec.target, connector.lastRows, connector.lastCols)
  }

  applyRangeSnapshot(snapshot: ConnectorRangeSnapshot, options: ConnectorAuthoritativeApplyOptions = {}): boolean {
    if (this.executionAuthority === 'server' && options.source !== 'collaboration') return false
    if (!isConnectorRangeSnapshot(snapshot)) return false
    const sheet = this.api.getActiveWorkbook()?.getSheetBySheetId(snapshot.sheetId)
    if (!sheet) return false
    sheet.getRange(snapshot.startRow, snapshot.startColumn, snapshot.rowCount, snapshot.columnCount).setValues(clone(snapshot.values))
    return true
  }

  captureRangeSnapshot(
    target: { sheetId: string; startRow: number; startColumn: number },
    rowCount: number,
    columnCount: number,
  ): ConnectorRangeSnapshot | null {
    if (!validDimension(target.startRow) || !validDimension(target.startColumn)
      || !validDimension(rowCount) || !validDimension(columnCount) || rowCount < 1 || columnCount < 1) return null
    return this.captureRange(target, rowCount, columnCount)
  }

  status(id: string): ConnectorStatus | undefined {
    const status = this.connectors.get(id)?.status
    return status ? { ...status } : undefined
  }

  serialize(): ConnectorSpec[] {
    return this.list()
  }

  async hydrate(specs: ConnectorSpec[]): Promise<void> {
    for (const spec of specs) this.add(spec)
    for (const spec of specs) if (spec.refresh === 'onOpen') await this.refresh(spec.id, { reason: 'onOpen' })
  }

  async refresh(id: string, options: ConnectorRefreshOptions = {}): Promise<ConnectorRefreshResult> {
    return this.runRefresh(id, options)
  }

  /** Refresh while capturing the complete affected cell rectangle. Failed or
   * partial writes are rolled back before the method returns. */
  async refreshTransaction(id: string, options: ConnectorRefreshOptions = {}): Promise<ConnectorRefreshTransaction> {
    const beforeState = this.snapshotState()
    let beforeRange: ConnectorRangeSnapshot | undefined
    let afterRange: ConnectorRangeSnapshot | undefined
    const result = await this.runRefresh(id, options, {
      beforeWrite: (connector, grid) => {
        const rowCount = Math.max(connector.lastRows, grid.length)
        const columnCount = Math.max(connector.lastCols, grid[0]?.length ?? 0)
        if (!rowCount || !columnCount) return
        const captured = this.captureRange(connector.spec.target, rowCount, columnCount)
        if (!captured) throw new Error('connector target sheet is unavailable')
        beforeRange = captured
      },
      afterWrite: () => {
        if (!beforeRange) return
        const captured = this.captureRange(beforeRange, beforeRange.rowCount, beforeRange.columnCount)
        if (!captured) throw new Error('connector target sheet is unavailable after write')
        afterRange = captured
      },
    })
    if (result !== 'applied' && !beforeRange) return { result, beforeState }
    if (result !== 'applied' || (!!beforeRange !== !!afterRange)) {
      let restored = true
      try {
        if (beforeRange) restored = this.applyRangeSnapshot(beforeRange)
      } finally {
        restored = this.restoreState(beforeState) && restored
      }
      if (!restored) throw new Error('connector refresh rollback could not restore the prior worksheet state')
      return { result: result === 'applied' ? 'failed' : result, beforeState }
    }
    return { result, beforeState, afterState: this.snapshotState(), beforeRange, afterRange }
  }

  private async runRefresh(
    id: string,
    options: ConnectorRefreshOptions,
    transaction?: {
      beforeWrite(connector: MountedConnector, grid: RangeValue[][]): void
      afterWrite(connector: MountedConnector, grid: RangeValue[][]): void
    },
  ): Promise<ConnectorRefreshResult> {
    if (this.executionAuthority === 'server') return 'skipped'
    const connector = this.connectors.get(id)
    if (!connector) return 'missing'
    if (connector.status.refreshing) return 'skipped'
    const reason = options.reason ?? 'manual'
    const controller = new AbortController()
    const relayAbort = () => controller.abort(options.signal?.reason)
    options.signal?.addEventListener('abort', relayAbort, { once: true })
    if (options.signal?.aborted) relayAbort()
    connector.controller = controller
    connector.status = { ...connector.status, refreshing: true, lastError: undefined, lastErrorCode: undefined, cacheHit: false }
    this.emitChange()
    try {
      if (this.authorize && !await this.authorize({ action: 'refresh', connector: structuredClone(connector.spec), reason })) return this.fail(connector, 'AUTHORIZATION_DENIED', 'refresh is not authorized')
      if (controller.signal.aborted) return this.fail(connector, 'ABORTED', 'refresh canceled')
      const cache = connector.spec.cache?.mode === 'memory' ? this.cache.get(id) : undefined
      let grid: RangeValue[][]
      if (!options.force && cache && cache.expiresAt >= this.now()) {
        grid = cloneGrid(cache.grid)
        connector.status.cacheHit = true
      } else {
        const fetched = await this.fetchSource(connector.spec.source, { connectorId: id, reason, signal: controller.signal })
        if (controller.signal.aborted) return this.fail(connector, 'ABORTED', 'refresh canceled')
        grid = normalizeGrid(fetched)
        if (this.preprocessing) {
          if (options.mode === 'collaborative' && !options.revision) return this.fail(connector, 'PREPROCESS_FAILED', 'collaborative preprocessing requires a source revision')
          const result = await this.preprocessing.run(grid, {
            range: {
              sheetId: connector.spec.target.sheetId,
              startRow: connector.spec.target.startRow,
              startColumn: connector.spec.target.startColumn,
              rowCount: grid.length,
              columnCount: grid[0]?.length ?? 0,
            },
            revision: options.revision ?? `local:${id}:${this.now()}`,
            mode: options.mode,
          }, { signal: controller.signal, expectedFingerprint: options.expectedPreprocessFingerprint })
          grid = result.grid
        }
        validateSchema(grid, connector.spec)
        if (connector.spec.cache?.mode === 'memory') this.cache.set(id, { expiresAt: this.now() + connector.spec.cache.ttlMs, grid: cloneGrid(grid) })
      }
      transaction?.beforeWrite(connector, grid)
      if (!this.writeGrid(connector, grid)) return this.fail(connector, 'TARGET_MISSING', 'target sheet is gone')
      transaction?.afterWrite(connector, grid)
      connector.status = { lastRefreshTs: this.now(), rowCount: grid.length, columnCount: grid[0]?.length ?? 0, refreshing: false, cacheHit: connector.status.cacheHit, nextRefreshTs: connector.status.nextRefreshTs }
      this.emitChange()
      return 'applied'
    } catch (error) {
      if (controller.signal.aborted || error instanceof RangePreprocessError && error.code === 'ABORTED') return this.fail(connector, 'ABORTED', 'refresh canceled')
      if (error instanceof RangePreprocessError) return this.fail(connector, 'PREPROCESS_FAILED', error.message)
      const code = error instanceof TypeError ? 'INVALID_DATA' : 'FETCH_FAILED'
      return this.fail(connector, code, error instanceof Error ? error.message : 'refresh failed')
    } finally {
      options.signal?.removeEventListener('abort', relayAbort)
      if (connector.controller === controller) connector.controller = undefined
      if (connector.status.refreshing) connector.status = { ...connector.status, refreshing: false }
      this.emitChange()
    }
  }

  private fail(connector: MountedConnector, code: NonNullable<ConnectorStatus['lastErrorCode']>, message: string): 'canceled' | 'failed' {
    connector.status = { ...connector.status, refreshing: false, lastError: message, lastErrorCode: code }
    return code === 'ABORTED' ? 'canceled' : 'failed'
  }

  private writeGrid(connector: MountedConnector, grid: RangeValue[][]): boolean {
    const sheet = this.api.getActiveWorkbook()?.getSheetBySheetId(connector.spec.target.sheetId)
    if (!sheet) return false
    const rows = grid.length
    const columns = grid[0]?.length ?? 0
    if (connector.lastRows > rows || connector.lastCols > columns) this.clearExtent(connector)
    if (rows && columns) {
      const cells: ICellData[][] = grid.map((row) => row.map((value) => ({ v: value })))
      sheet.getRange(connector.spec.target.startRow, connector.spec.target.startColumn, rows, columns).setValues(cells)
    }
    connector.lastRows = rows
    connector.lastCols = columns
    return true
  }

  private clearExtent(connector: MountedConnector): void {
    if (!connector.lastRows || !connector.lastCols) return
    const sheet = this.api.getActiveWorkbook()?.getSheetBySheetId(connector.spec.target.sheetId)
    if (!sheet) return
    const blanks: ICellData[][] = Array.from({ length: connector.lastRows }, () => Array.from({ length: connector.lastCols }, () => ({ v: null })))
    sheet.getRange(connector.spec.target.startRow, connector.spec.target.startColumn, connector.lastRows, connector.lastCols).setValues(blanks)
  }

  private captureRange(
    target: { sheetId: string; startRow: number; startColumn: number },
    rowCount: number,
    columnCount: number,
  ): ConnectorRangeSnapshot | null {
    const sheet = this.api.getActiveWorkbook()?.getSheetBySheetId(target.sheetId)
    if (!sheet) return null
    const range = sheet.getRange(target.startRow, target.startColumn, rowCount, columnCount)
    const values = range.getValues()
    const formulas = (range as typeof range & { getFormulas?: () => Array<Array<string | null | undefined>> }).getFormulas?.()
    const cells: ICellData[][] = values.map((row, rowIndex) => row.map((value, columnIndex) => {
      const formula = formulas?.[rowIndex]?.[columnIndex]
      return typeof formula === 'string' && formula.length > 0 ? { f: formula } : { v: value ?? null }
    }))
    return {
      sheetId: target.sheetId,
      startRow: target.startRow,
      startColumn: target.startColumn,
      rowCount,
      columnCount,
      values: clone(cells),
    }
  }

  private startSchedule(connector: MountedConnector): void {
    if (this.executionAuthority === 'server') return
    if (connector.spec.refresh !== 'interval' || !connector.spec.schedule) return
    const intervalMs = connector.spec.schedule.intervalMs
    connector.status.nextRefreshTs = this.now() + intervalMs
    connector.timer = this.scheduler.setInterval(() => {
      connector.status.nextRefreshTs = this.now() + intervalMs
      void this.refresh(connector.spec.id, { reason: 'interval' })
    }, intervalMs)
  }

  private stop(connector: MountedConnector): void {
    connector.controller?.abort('connector removed')
    if (connector.timer !== undefined) this.scheduler.clearInterval(connector.timer)
    connector.timer = undefined
  }
}
