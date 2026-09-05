// Stage B — cell-level co-editing over the gateway's server-ordered op log.
//
// Pure pieces (tested without Univer): which local mutations to ship, how
// to reconcile an incoming op against our own later edits so every client
// converges on "last writer by seq wins" per cell, and the submit queue.
// The Univer-facing glue lives in manager.ts.

/** One Univer mutation on the wire. params is a deep-cloned plain object. */
export interface OpRecord {
  id: string
  params: Record<string, unknown>
}

/** One ordered entry of the room log (gateway OpEntry). */
export interface OpEntry {
  room: string
  seq: number
  client_id: string
  ops: OpRecord[]
}

export interface LogState {
  seq: number
  saved_seq: number
}

/** What the Univer command service hands a mutation listener. */
export interface MutationInfo {
  id: string
  type?: number
  params?: unknown
}

export interface ExecOptions {
  onlyLocal?: boolean
  fromCollab?: boolean
  fromChangeset?: boolean
  syncOnly?: boolean
  [k: string]: unknown
}

export const SET_RANGE_VALUES = 'sheet.mutation.set-range-values'
export const MOVE_RANGE = 'sheet.mutation.move-range'
export const REORDER_RANGE = 'sheet.mutation.reorder-range'

/**
 * Mutations never shipped even though they target the sheet. Freeze panes
 * are kept local by Univer itself (it forces onlyLocal and self-heals).
 * Every other sheet.mutation.* (numfmt, merges, sizes, sheet add/rename,
 * and structural row/column edits) ships: it's data. Stage C transforms the
 * supported insert/remove/move vocabulary before queued operations resubmit.
 */
const NEVER_SHIP = new Set<string>(['sheet.mutation.set-frozen'])

/**
 * Decide whether a locally executed mutation goes on the wire.
 *   - only MUTATIONs (type 2) with a sheet.mutation.* id — doc mutations are
 *     the cell editor's scratch document, formula.* are engine internals
 *   - never echoes: fromCollab/fromChangeset (remote or snapshot replay)
 *   - never onlyLocal/syncOnly: that is how Univer keeps formula results,
 *     freeze and other derived state machine-local
 *   - only our workbook (unitId), so a second unit on the page stays out
 */
export function shouldShipMutation(info: MutationInfo, options: ExecOptions | undefined, unitId: string): boolean {
  if (options?.fromCollab || options?.fromChangeset || options?.onlyLocal || options?.syncOnly) return false
  if (info.type !== undefined && info.type !== 2) return false
  if (!info.id.startsWith('sheet.mutation.')) return false
  if (NEVER_SHIP.has(info.id)) return false
  const p = info.params as { unitId?: unknown } | undefined
  if (!p || typeof p !== 'object') return false
  if (p.unitId !== undefined && p.unitId !== unitId) return false
  return true
}

/**
 * Rebase ops onto this client's workbook id. Univer assigns a fresh unitId
 * on every createWorkbook, so the same file has a different id in every
 * tab; sheet ids (subUnitId) come from the import and are stable. Without
 * this a replayed mutation targets a unit the receiver doesn't have and is
 * silently a no-op — and reconcile's cell keys (which include the unit)
 * never match. Applied before anything else touches a remote op.
 */
export function rebaseUnit(ops: OpRecord[], unitId: string): OpRecord[] {
  return ops.map((op) => {
    const p = op.params
    if (!p || typeof p !== 'object' || p.unitId === undefined || p.unitId === unitId) return op
    return { id: op.id, params: { ...p, unitId } }
  })
}

/** Plain deep copy — the command service mutates params in place (trigger stamping). */
export function cloneParams(params: unknown): Record<string, unknown> {
  return JSON.parse(JSON.stringify(params ?? {})) as Record<string, unknown>
}

type CellMatrix = Record<string, Record<string, unknown>>

interface SetRangeValuesParams {
  unitId?: string
  subUnitId?: string
  cellValue?: CellMatrix | null
}

interface RangeLike {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}

function cellKey(unit: string, sub: string, r: string, c: string): string {
  return `${unit} ${sub} ${r} ${c}`
}

function matrixCells(out: Set<string>, unit: string, sub: string, matrix: CellMatrix | null | undefined): void {
  if (!matrix) return
  for (const row of Object.keys(matrix)) {
    const columns = matrix[row]
    if (!columns) continue
    for (const column of Object.keys(columns)) out.add(cellKey(unit, sub, row, column))
  }
}

function validRange(value: unknown): value is RangeLike {
  if (!value || typeof value !== 'object') return false
  const range = value as Record<string, unknown>
  return [range.startRow, range.endRow, range.startColumn, range.endColumn].every(Number.isSafeInteger)
    && (range.startRow as number) >= 0
    && (range.startColumn as number) >= 0
    && (range.endRow as number) >= (range.startRow as number)
    && (range.endColumn as number) >= (range.startColumn as number)
}

/** Cells written by set-values, move-range matrices, or a row reorder. */
export function cellsOf(op: OpRecord): Set<string> {
  const out = new Set<string>()
  const unit = String(op.params.unitId ?? '')
  if (op.id === SET_RANGE_VALUES) {
    const params = op.params as SetRangeValuesParams
    matrixCells(out, unit, String(params.subUnitId ?? ''), params.cellValue)
  } else if (op.id === MOVE_RANGE) {
    const params = op.params as { from?: { subUnitId?: unknown; value?: CellMatrix }; to?: { subUnitId?: unknown; value?: CellMatrix } }
    for (const part of [params.from, params.to]) if (part) matrixCells(out, unit, String(part.subUnitId ?? ''), part.value)
  } else if (op.id === REORDER_RANGE) {
    const params = op.params as { subUnitId?: unknown; range?: unknown; order?: Record<string, unknown> }
    if (validRange(params.range) && params.order && typeof params.order === 'object') {
      const sub = String(params.subUnitId ?? '')
      for (let row = params.range.startRow; row <= params.range.endRow; row++) {
        if (!Object.prototype.hasOwnProperty.call(params.order, row)) continue
        for (let column = params.range.startColumn; column <= params.range.endColumn; column++) out.add(cellKey(unit, sub, String(row), String(column)))
      }
    }
  }
  return out
}

function filterMatrix(matrix: CellMatrix | undefined, unit: string, sub: string, protect: ReadonlySet<string>): CellMatrix {
  const kept: CellMatrix = {}
  for (const [row, columns] of Object.entries(matrix ?? {})) {
    for (const [column, value] of Object.entries(columns ?? {})) {
      if (protect.has(cellKey(unit, sub, row, column))) continue
      const target = kept[row] ?? (kept[row] = {})
      target[column] = value
    }
  }
  return kept
}

function splitReorder(op: OpRecord, protect: ReadonlySet<string>): OpRecord[] {
  const params = op.params as Record<string, unknown> & { subUnitId?: unknown; range?: unknown; order?: Record<string, unknown> }
  if (!validRange(params.range) || !params.order || typeof params.order !== 'object') return [op]
  const unit = String(params.unitId ?? '')
  const sub = String(params.subUnitId ?? '')
  const groups: Array<{ start: number; end: number; signature: string; order: Record<string, unknown> }> = []
  for (let column = params.range.startColumn; column <= params.range.endColumn; column++) {
    const order: Record<string, unknown> = {}
    for (let row = params.range.startRow; row <= params.range.endRow; row++) {
      if (!Object.prototype.hasOwnProperty.call(params.order, row)) continue
      const sourceRow = params.order[String(row)]
      if (!Number.isSafeInteger(sourceRow)) continue
      // Reorder reads from the live sheet. If either destination or source is
      // locally owned, applying it could overwrite or copy a later local value.
      if (!protect.has(cellKey(unit, sub, String(row), String(column))) && !protect.has(cellKey(unit, sub, String(sourceRow), String(column)))) order[String(row)] = sourceRow
    }
    const signature = JSON.stringify(order)
    if (signature === '{}') continue
    const previous = groups.at(-1)
    if (previous && previous.end + 1 === column && previous.signature === signature) previous.end = column
    else groups.push({ start: column, end: column, signature, order })
  }
  return groups.map((group) => ({
    id: op.id,
    params: { ...params, range: { ...params.range as RangeLike, startColumn: group.start, endColumn: group.end }, order: group.order },
  }))
}

/** Remove later-owned cells from a content mutation without mutating it. */
export function withoutCells(op: OpRecord, protect: ReadonlySet<string>): OpRecord[] {
  if (protect.size === 0) return [op]
  const unit = String(op.params.unitId ?? '')
  if (op.id === SET_RANGE_VALUES) {
    const params = op.params as SetRangeValuesParams
    const cellValue = filterMatrix(params.cellValue ?? undefined, unit, String(params.subUnitId ?? ''), protect)
    return Object.keys(cellValue).length ? [{ id: op.id, params: { ...op.params, cellValue } }] : []
  }
  if (op.id === MOVE_RANGE) {
    const params = op.params as Record<string, unknown> & { from?: { subUnitId?: unknown; value?: CellMatrix }; to?: { subUnitId?: unknown; value?: CellMatrix } }
    const from = params.from ? { ...params.from, value: filterMatrix(params.from.value, unit, String(params.from.subUnitId ?? ''), protect) } : undefined
    const to = params.to ? { ...params.to, value: filterMatrix(params.to.value, unit, String(params.to.subUnitId ?? ''), protect) } : undefined
    if (!Object.keys(from?.value ?? {}).length && !Object.keys(to?.value ?? {}).length) return []
    return [{ id: op.id, params: { ...params, ...(from ? { from } : {}), ...(to ? { to } : {}) } }]
  }
  if (op.id === REORDER_RANGE) return splitReorder(op, protect)
  return [op]
}

/** A local submission we still consider "after" incoming remote ops. */
export interface LocalOwned {
  /** Assigned seq, or null while unacknowledged (it will be later than anything in flight). */
  seq: number | null
  cells: Set<string>
}

/**
 * Reconcile an incoming remote entry against our own later edits: any cell
 * we wrote in an op that is pending or has a higher seq is OURS (we are the
 * last writer by server order), so it is dropped from the remote op before
 * it is applied. Every client does this, so all converge. Returns the ops to
 * apply (a set-range-values with nothing left is dropped).
 */
export function reconcileRemote(entry: OpEntry, owned: LocalOwned[]): OpRecord[] {
  const protect = new Set<string>()
  for (const o of owned) {
    if (o.seq === null || o.seq > entry.seq) for (const k of o.cells) protect.add(k)
  }
  if (protect.size === 0) return entry.ops
  const out: OpRecord[] = []
  for (const op of entry.ops) out.push(...withoutCells(op, protect))
  return out
}

/** Forget owned entries that no incoming op can be older than any more. */
export function pruneOwned(owned: LocalOwned[], remoteSeq: number): LocalOwned[] {
  return owned.filter((o) => o.seq === null || o.seq > remoteSeq)
}

/**
 * SubmitQueue — batches mutations emitted in one tick into a single
 * submission and ships submissions strictly in order, one in flight at a
 * time. A failed submission (socket down) stays queued and is retried on
 * flush(); the host calls flush() after a reconnect + catch-up.
 */
export class SubmitQueue {
  private batch: OpRecord[] = []
  private queued: OpRecord[][] = []
  private inFlight = false
  private scheduled = false
  private readonly send: (ops: OpRecord[]) => Promise<number>
  private readonly onAcked: (ops: OpRecord[], seq: number) => void
  private readonly onQueued: (ops: OpRecord[]) => void

  private readonly onStale?: () => Promise<void>

  constructor(
    send: (ops: OpRecord[]) => Promise<number>,
    hooks: {
      onQueued: (ops: OpRecord[]) => void
      onAcked: (ops: OpRecord[], seq: number) => void
      /** Called when the gateway rejects a submission as STALE_BASE; catch up, then the queue retries. */
      onStale?: () => Promise<void>
    },
  ) {
    this.send = send
    this.onQueued = hooks.onQueued
    this.onAcked = hooks.onAcked
    this.onStale = hooks.onStale
  }

  /** Everything not yet acknowledged, oldest first (current batch last). */
  allPending(): OpRecord[] {
    return [...this.queued.flat(), ...this.batch]
  }

  /**
   * Rewrite every unacknowledged op (Stage C: transform pending through an
   * incoming entry). Submissions left empty are dropped.
   */
  transformQueued(fn: (ops: OpRecord[]) => OpRecord[]): void {
    this.batch = this.batch.length ? fn(this.batch) : this.batch
    this.queued = this.queued.map((ops) => fn(ops)).filter((ops) => ops.length > 0)
  }

  /** Queue one mutation; everything pushed in the same microtask ships together. */
  push(op: OpRecord): void {
    this.batch.push(op)
    if (this.scheduled) return
    this.scheduled = true
    queueMicrotask(() => {
      this.scheduled = false
      if (this.batch.length === 0) return
      const ops = this.batch
      this.batch = []
      this.queued.push(ops)
      this.onQueued(ops)
      void this.flush()
    })
  }

  get pending(): number {
    return this.queued.length + (this.batch.length ? 1 : 0)
  }

  /** Ship queued submissions in order; stops at the first failure (retried on the next flush). */
  async flush(): Promise<void> {
    if (this.inFlight) return
    this.inFlight = true
    try {
      while (this.queued.length > 0) {
        // Re-read each attempt: a stale-base catch-up may have transformed
        // (or emptied) the head submission while we were waiting.
        const ops = this.queued[0]
        let seq: number
        try {
          seq = await this.send(ops)
        } catch (e) {
          if (this.onStale && e instanceof Error && e.message.includes('STALE_BASE')) {
            try {
              await this.onStale()
            } catch {
              return
            }
            continue
          }
          return
        }
        this.queued.shift()
        this.onAcked(ops, seq)
      }
    } finally {
      this.inFlight = false
    }
  }

  clear(): void {
    this.batch = []
    this.queued = []
  }
}
