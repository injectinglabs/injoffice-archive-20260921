import type { FileChartAnchor } from './fromFile'
import { ChartCommandController, isChartSnapshotV1 } from './commands'
import type { ChartLayerOperation, ChartSpec } from './types'

const PROTOCOL = 'injoffice.chart-collaboration/v1' as const
const LAYER_OPERATIONS: readonly ChartLayerOperation[] = ['bringForward', 'sendBackward', 'bringToFront', 'sendToBack']

export type ChartUpdatePatch = Partial<Omit<ChartSpec, 'id' | 'nativeIdentity'>>

export type ChartCollaborationMutation =
  | { kind: 'create'; chart: ChartSpec; cellAnchor?: FileChartAnchor; baseObjectRevision: 0; baseLayerRevision: number }
  | { kind: 'update'; chartId: string; patch: ChartUpdatePatch; baseObjectRevision: number }
  | { kind: 'remove'; chartId: string; sheetId: string; baseObjectRevision: number; baseLayerRevision: number }
  | { kind: 'layer'; chartId: string; sheetId: string; operation: ChartLayerOperation; baseObjectRevision: number; baseLayerRevision: number }

export interface ChartCollaborationIntent {
  protocol: typeof PROTOCOL
  workbookId: string
  operationId: string
  actorId: string
  authorityRevision: number
  mutation: ChartCollaborationMutation
}

/** An intent after a collaboration authority assigns its total-order revision. */
export interface ChartCollaborationEvent extends ChartCollaborationIntent {
  revision: number
}

export interface ChartCollaborationAuthorityRequest {
  workbookId: string
  operationId: string
  actorId: string
  revision: number
  mutation: Readonly<ChartCollaborationMutation>
}

/** A synchronous authority hook. Its revision must change whenever permissions do. */
export interface ChartCollaborationAuthority {
  readonly revision: number
  canApply(request: ChartCollaborationAuthorityRequest): boolean
}

export type ChartCollaborationResultCode =
  | 'applied'
  | 'no-op'
  | 'malformed'
  | 'wrong-workbook'
  | 'stale'
  | 'gap'
  | 'desynchronized'
  | 'authority-revision'
  | 'authority-failed'
  | 'unauthorized'
  | 'conflict'
  | 'application-failed'

export interface ChartCollaborationResult {
  ok: boolean
  applied: boolean
  consumed: boolean
  code: ChartCollaborationResultCode
  revision: number
  operationId?: string
}

export interface ChartCollaborationObjectRevision {
  chartId: string
  sheetId: string
  revision: number
  deleted: boolean
}

export interface ChartCollaborationLayerRevision {
  sheetId: string
  revision: number
}

/** Revision metadata only. Chart content is hydrated through ChartManager. */
export interface ChartCollaborationRevisionStateV1 {
  version: 1
  revision: number
  objects: ChartCollaborationObjectRevision[]
  layers: ChartCollaborationLayerRevision[]
}

export interface ChartCollaborationSessionOptions {
  authority: ChartCollaborationAuthority
  operationId?: () => string
  onResyncRequired?: (result: ChartCollaborationResult) => void
}

interface ObjectRevision {
  sheetId: string
  revision: number
  deleted: boolean
}

let operationSequence = 0

function defaultOperationId(): string {
  return `chart-op-${Date.now().toString(36)}-${(++operationSequence).toString(36)}`
}

function validId(value: unknown): value is string {
  return typeof value === 'string' && value.length > 0 && value.length <= 512
}

function validRevision(value: unknown): value is number {
  return Number.isSafeInteger(value) && (value as number) >= 0
}

function isJsonValue(value: unknown, ancestors = new Set<object>()): boolean {
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return true
  if (typeof value === 'number') return Number.isFinite(value)
  if (typeof value !== 'object') return false
  if (ancestors.has(value)) return false
  const prototype = Object.getPrototypeOf(value)
  if (prototype !== Object.prototype && prototype !== Array.prototype) return false
  ancestors.add(value)
  const valid = Array.isArray(value)
    ? value.every((entry) => isJsonValue(entry, ancestors))
    : Object.entries(value).every(([key, entry]) => key.length <= 512 && isJsonValue(entry, ancestors))
  ancestors.delete(value)
  return valid
}

function copy<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T
}

function compareId(left: string, right: string): number {
  return left < right ? -1 : left > right ? 1 : 0
}

function hasOwn(value: object, key: string): boolean {
  return Object.prototype.hasOwnProperty.call(value, key)
}

function validLayerOperation(value: unknown): value is ChartLayerOperation {
  return LAYER_OPERATIONS.includes(value as ChartLayerOperation)
}

function validPatch(value: unknown): value is ChartUpdatePatch {
  return !!value && typeof value === 'object' && !Array.isArray(value) && isJsonValue(value)
    && !hasOwn(value, 'id') && !hasOwn(value, 'nativeIdentity')
}

function validMutation(value: unknown): value is ChartCollaborationMutation {
  if (!value || typeof value !== 'object' || !isJsonValue(value)) return false
  const mutation = value as Partial<ChartCollaborationMutation> & Record<string, unknown>
  if (mutation.kind === 'create') {
    return mutation.baseObjectRevision === 0 && validRevision(mutation.baseLayerRevision)
      && isChartSnapshotV1({ version: 1, charts: [{ spec: mutation.chart, cellAnchor: mutation.cellAnchor }] })
  }
  if (mutation.kind === 'update') {
    return validId(mutation.chartId) && validRevision(mutation.baseObjectRevision) && validPatch(mutation.patch)
  }
  if (mutation.kind === 'remove') {
    return validId(mutation.chartId) && validId(mutation.sheetId)
      && validRevision(mutation.baseObjectRevision) && validRevision(mutation.baseLayerRevision)
  }
  if (mutation.kind === 'layer') {
    return validId(mutation.chartId) && validId(mutation.sheetId) && validLayerOperation(mutation.operation)
      && validRevision(mutation.baseObjectRevision) && validRevision(mutation.baseLayerRevision)
  }
  return false
}

function validEnvelope(value: unknown): value is ChartCollaborationEvent {
  if (!value || typeof value !== 'object' || !isJsonValue(value)) return false
  const event = value as Partial<ChartCollaborationEvent>
  return event.protocol === PROTOCOL && validId(event.workbookId) && validId(event.operationId)
    && validId(event.actorId) && validRevision(event.authorityRevision)
    && validRevision(event.revision) && event.revision > 0 && validMutation(event.mutation)
}

/**
 * Applies server-ordered chart-object events without ever restoring a complete
 * chart collection. Object revisions serialize edits to one stable chart id;
 * sheet-layer revisions serialize the relational back-to-front order.
 */
export class ChartCollaborationSession {
  private revisionValue = 0
  private desynchronizedValue = false
  private readonly objects = new Map<string, ObjectRevision>()
  private readonly layers = new Map<string, number>()
  private readonly listeners = new Set<(result: ChartCollaborationResult, event: ChartCollaborationEvent) => void>()
  private readonly operationId: () => string

  constructor(
    readonly workbookId: string,
    readonly commands: ChartCommandController,
    private readonly options: ChartCollaborationSessionOptions,
  ) {
    if (!validId(workbookId)) throw new TypeError('workbookId must be a non-empty string')
    this.operationId = options.operationId ?? defaultOperationId
    for (const chart of commands.manager.list()) {
      this.objects.set(chart.id, { sheetId: chart.range.sheetId, revision: 0, deleted: false })
      if (!this.layers.has(chart.range.sheetId)) this.layers.set(chart.range.sheetId, 0)
    }
  }

  get revision(): number {
    return this.revisionValue
  }

  get desynchronized(): boolean {
    return this.desynchronizedValue
  }

  onResult(listener: (result: ChartCollaborationResult, event: ChartCollaborationEvent) => void): () => void {
    this.listeners.add(listener)
    return () => this.listeners.delete(listener)
  }

  createIntent(actorId: string, chart: ChartSpec, cellAnchor?: FileChartAnchor): ChartCollaborationIntent | null {
    if (!this.canBuild(actorId) || this.objects.has(chart.id) || this.commands.manager.getSpec(chart.id)) return null
    const mutation: ChartCollaborationMutation = {
      kind: 'create', chart: copy(chart), ...(cellAnchor ? { cellAnchor: copy(cellAnchor) } : {}),
      baseObjectRevision: 0, baseLayerRevision: this.layers.get(chart.range.sheetId) ?? 0,
    }
    return this.intent(actorId, mutation)
  }

  updateIntent(actorId: string, chartId: string, patch: ChartUpdatePatch): ChartCollaborationIntent | null {
    const state = this.liveObject(chartId)
    const current = this.commands.manager.getSpec(chartId)
    if (!this.canBuild(actorId) || !state || !current || !validPatch(patch)) return null
    const candidate: ChartSpec = { ...current, ...patch, id: chartId, nativeIdentity: current.nativeIdentity }
    if (candidate.range.sheetId !== state.sheetId || !isChartSnapshotV1({ version: 1, charts: [{ spec: candidate }] })) return null
    return this.intent(actorId, { kind: 'update', chartId, patch: copy(patch), baseObjectRevision: state.revision })
  }

  removeIntent(actorId: string, chartId: string): ChartCollaborationIntent | null {
    const state = this.liveObject(chartId)
    if (!this.canBuild(actorId) || !state || !this.commands.manager.getSpec(chartId)) return null
    return this.intent(actorId, {
      kind: 'remove', chartId, sheetId: state.sheetId, baseObjectRevision: state.revision,
      baseLayerRevision: this.layers.get(state.sheetId) ?? 0,
    })
  }

  layerIntent(actorId: string, chartId: string, operation: ChartLayerOperation): ChartCollaborationIntent | null {
    const state = this.liveObject(chartId)
    if (!this.canBuild(actorId) || !state || !validLayerOperation(operation) || !this.commands.manager.canLayer(chartId, operation)) return null
    return this.intent(actorId, {
      kind: 'layer', chartId, sheetId: state.sheetId, operation, baseObjectRevision: state.revision,
      baseLayerRevision: this.layers.get(state.sheetId) ?? 0,
    })
  }

  receive(untrustedEvent: unknown): ChartCollaborationResult {
    if (!validEnvelope(untrustedEvent)) return this.result(false, false, false, 'malformed')
    const event = copy(untrustedEvent)
    if (event.workbookId !== this.workbookId) return this.result(false, false, false, 'wrong-workbook', event.operationId)
    if (this.desynchronizedValue) return this.result(false, false, false, 'desynchronized', event.operationId)
    if (event.revision <= this.revisionValue) return this.result(false, false, false, 'stale', event.operationId)
    if (event.revision !== this.revisionValue + 1) return this.failClosed('gap', event)

    let authorityRevision: number
    try {
      authorityRevision = this.options.authority.revision
    } catch {
      return this.failClosed('authority-failed', event)
    }
    if (!validRevision(authorityRevision) || event.authorityRevision !== authorityRevision) {
      return this.failClosed('authority-revision', event)
    }

    let authorized: boolean
    try {
      authorized = this.options.authority.canApply({
        workbookId: event.workbookId,
        operationId: event.operationId,
        actorId: event.actorId,
        revision: event.revision,
        mutation: copy(event.mutation),
      })
    } catch {
      return this.failClosed('authority-failed', event)
    }
    if (authorized !== true) return this.consume(event, false, false, 'unauthorized')

    return this.apply(event)
  }

  revisionState(): ChartCollaborationRevisionStateV1 {
    return {
      version: 1,
      revision: this.revisionValue,
      objects: [...this.objects].map(([chartId, state]) => ({ chartId, ...state })).sort((a, b) => compareId(a.chartId, b.chartId)),
      layers: [...this.layers].map(([sheetId, revision]) => ({ sheetId, revision })).sort((a, b) => compareId(a.sheetId, b.sheetId)),
    }
  }

  /** Replace only collaboration revision metadata after the host has hydrated
   * the matching chart collection through ChartManager. */
  resync(state: ChartCollaborationRevisionStateV1): boolean {
    if (!this.validRevisionState(state)) return false
    this.revisionValue = state.revision
    this.objects.clear()
    this.layers.clear()
    for (const object of state.objects) this.objects.set(object.chartId, {
      sheetId: object.sheetId, revision: object.revision, deleted: object.deleted,
    })
    for (const layer of state.layers) this.layers.set(layer.sheetId, layer.revision)
    this.desynchronizedValue = false
    return true
  }

  private intent(actorId: string, mutation: ChartCollaborationMutation): ChartCollaborationIntent | null {
    let authorityRevision: number
    let operationId: string
    try {
      authorityRevision = this.options.authority.revision
      operationId = this.operationId()
    } catch {
      return null
    }
    if (!validRevision(authorityRevision) || !validId(operationId) || !validMutation(mutation)) return null
    return { protocol: PROTOCOL, workbookId: this.workbookId, operationId, actorId, authorityRevision, mutation }
  }

  private canBuild(actorId: string): boolean {
    return !this.desynchronizedValue && validId(actorId)
  }

  private liveObject(chartId: string): ObjectRevision | undefined {
    const state = this.objects.get(chartId)
    return state && !state.deleted ? state : undefined
  }

  private apply(event: ChartCollaborationEvent): ChartCollaborationResult {
    const mutation = event.mutation
    if (mutation.kind === 'create') {
      if (this.objects.has(mutation.chart.id) || this.commands.manager.getSpec(mutation.chart.id)
        || mutation.baseLayerRevision !== (this.layers.get(mutation.chart.range.sheetId) ?? 0)) {
        return this.consume(event, false, false, 'conflict')
      }
      if (!this.applyExternal(() => this.commands.applyExternalCreate(mutation.chart, mutation.cellAnchor))) return this.failClosed('application-failed', event)
      this.objects.set(mutation.chart.id, { sheetId: mutation.chart.range.sheetId, revision: event.revision, deleted: false })
      this.layers.set(mutation.chart.range.sheetId, event.revision)
      return this.consume(event, true, true, 'applied')
    }

    const state = this.liveObject(mutation.chartId)
    if (!state || !this.commands.manager.getSpec(mutation.chartId) || mutation.baseObjectRevision !== state.revision) {
      return this.consume(event, false, false, 'conflict')
    }

    if (mutation.kind === 'update') {
      const current = this.commands.manager.getSpec(mutation.chartId)!
      const candidate: ChartSpec = { ...current, ...mutation.patch, id: mutation.chartId, nativeIdentity: current.nativeIdentity }
      if (candidate.range.sheetId !== state.sheetId || !isChartSnapshotV1({ version: 1, charts: [{ spec: candidate }] })) {
        return this.consume(event, false, false, 'conflict')
      }
      if (!this.applyExternal(() => this.commands.applyExternalUpdate(mutation.chartId, mutation.patch))) return this.failClosed('application-failed', event)
      state.revision = event.revision
      return this.consume(event, true, true, 'applied')
    }

    if (mutation.sheetId !== state.sheetId || mutation.baseLayerRevision !== (this.layers.get(state.sheetId) ?? 0)) {
      return this.consume(event, false, false, 'conflict')
    }

    if (mutation.kind === 'remove') {
      if (!this.applyExternal(() => this.commands.applyExternalRemove(mutation.chartId))) return this.failClosed('application-failed', event)
      this.objects.set(mutation.chartId, { ...state, revision: event.revision, deleted: true })
      this.layers.set(state.sheetId, event.revision)
      return this.consume(event, true, true, 'applied')
    }

    if (!this.commands.manager.canLayer(mutation.chartId, mutation.operation)) {
      state.revision = event.revision
      this.layers.set(state.sheetId, event.revision)
      return this.consume(event, false, true, 'no-op')
    }
    if (!this.applyExternal(() => this.commands.applyExternalLayer(mutation.chartId, mutation.operation))) return this.failClosed('application-failed', event)
    state.revision = event.revision
    this.layers.set(state.sheetId, event.revision)
    return this.consume(event, true, true, 'applied')
  }

  private consume(
    event: ChartCollaborationEvent,
    applied: boolean,
    ok: boolean,
    code: ChartCollaborationResultCode,
  ): ChartCollaborationResult {
    this.revisionValue = event.revision
    const result = this.result(ok, applied, true, code, event.operationId)
    this.notify(result, event)
    return result
  }

  private applyExternal(application: () => boolean): boolean {
    try {
      return application() === true
    } catch {
      return false
    }
  }

  private failClosed(code: 'gap' | 'authority-revision' | 'authority-failed' | 'application-failed', event: ChartCollaborationEvent): ChartCollaborationResult {
    this.desynchronizedValue = true
    const result = this.result(false, false, false, code, event.operationId)
    try {
      this.options.onResyncRequired?.({ ...result })
    } catch {
      // Observability must not change collaboration state or outcomes.
    }
    this.notify(result, event)
    return result
  }

  private notify(result: ChartCollaborationResult, event: ChartCollaborationEvent): void {
    for (const listener of this.listeners) {
      try {
        listener({ ...result }, copy(event))
      } catch {
        // One observer cannot prevent other observers or roll back an event.
      }
    }
  }

  private result(
    ok: boolean,
    applied: boolean,
    consumed: boolean,
    code: ChartCollaborationResultCode,
    operationId?: string,
  ): ChartCollaborationResult {
    return { ok, applied, consumed, code, revision: this.revisionValue, ...(operationId ? { operationId } : {}) }
  }

  private validRevisionState(value: unknown): value is ChartCollaborationRevisionStateV1 {
    if (!value || typeof value !== 'object' || !isJsonValue(value)) return false
    const state = value as Partial<ChartCollaborationRevisionStateV1>
    if (state.version !== 1 || !validRevision(state.revision) || !Array.isArray(state.objects) || !Array.isArray(state.layers)) return false
    const objects = new Map<string, ChartCollaborationObjectRevision>()
    for (const object of state.objects) {
      if (!object || !validId(object.chartId) || !validId(object.sheetId) || !validRevision(object.revision)
        || object.revision > state.revision || typeof object.deleted !== 'boolean' || objects.has(object.chartId)) return false
      objects.set(object.chartId, object)
    }
    const layers = new Map<string, number>()
    for (const layer of state.layers) {
      if (!layer || !validId(layer.sheetId) || !validRevision(layer.revision)
        || layer.revision > state.revision || layers.has(layer.sheetId)) return false
      layers.set(layer.sheetId, layer.revision)
    }
    for (const chart of this.commands.manager.list()) {
      const object = objects.get(chart.id)
      if (!object || object.deleted || object.sheetId !== chart.range.sheetId || !layers.has(chart.range.sheetId)) return false
    }
    for (const object of objects.values()) {
      const chart = this.commands.manager.getSpec(object.chartId)
      if (object.deleted ? !!chart : !chart) return false
    }
    return true
  }
}
