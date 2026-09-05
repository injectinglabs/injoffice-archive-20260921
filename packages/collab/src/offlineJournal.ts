import type { OpRecord } from './sync'

export const OUTBOUND_JOURNAL_PROTOCOL = 'injoffice.collab.outbound-journal/v1' as const

export interface OutboundJournalEntry {
  idempotencyKey: string
  ordinal: number
  createdBaseSeq: number
  attempts: number
  ops: OpRecord[]
}

export interface OutboundJournalReceipt {
  idempotencyKey: string
  entry: OutboundJournalEntry
  acknowledged: Promise<number>
  /** Release a durably enqueued entry for transport. Idempotent. */
  release(): void
}

export interface OutboundJournalEnqueueOptions {
  /** Keep this entry from reaching the transport until receipt.release(). */
  deferSend?: boolean
}

export interface OutboundJournalSnapshotV1 {
  protocol: typeof OUTBOUND_JOURNAL_PROTOCOL
  room: string
  clientId: string
  lastServerSeq: number
  nextOrdinal: number
  entries: OutboundJournalEntry[]
}

export interface OutboundJournalStorage {
  read(key: string): string | null | Promise<string | null>
  write(key: string, value: string): void | Promise<void>
}

export interface RetryScheduler {
  schedule(delayMs: number, task: () => void): { cancel(): void }
}

export interface JournalSubmitRequest {
  room: string
  clientId: string
  idempotencyKey: string
  baseSeq: number
  ops: OpRecord[]
}

export interface JournalSubmitAck {
  idempotencyKey: string
  seq: number
  /** Required when returning an earlier/already-observed sequence. */
  deduplicated?: boolean
}

export interface JournalResyncRequest {
  room: string
  clientId: string
  afterSeq: number
  pending: OutboundJournalEntry[]
}

export interface JournalRebasedEntry {
  idempotencyKey: string
  ops: OpRecord[]
}

export interface JournalResyncResult {
  headSeq: number
  replay: 'safe' | 'blocked'
  /** If supplied, must contain exactly one rewrite for every pending key. */
  rebased?: JournalRebasedEntry[]
}

export interface OutboundJournalTransport {
  submit(request: JournalSubmitRequest, signal: AbortSignal): Promise<JournalSubmitAck>
  /** Host catches up remote state here before local entries can replay. */
  resync(request: JournalResyncRequest, signal: AbortSignal): Promise<JournalResyncResult>
}

export interface OutboundRetryPolicy {
  initialDelayMs: number
  multiplier: number
  maxDelayMs: number
}

export type OutboundJournalStatus =
  | 'closed'
  | 'offline'
  | 'resyncing'
  | 'ready'
  | 'sending'
  | 'retry-wait'
  | 'blocked'
  | 'disposed'

export interface OutboundJournalState {
  status: OutboundJournalStatus
  pending: number
  lastServerSeq: number
  retryAttempt: number
  nextRetryMs: number | null
}

export interface OutboundJournalHooks {
  onState?: (state: OutboundJournalState) => void
  onAck?: (entry: OutboundJournalEntry, ack: JournalSubmitAck) => void
  onBlocked?: (reason: 'unsafe-replay' | 'invalid-resync', detail?: unknown) => void
}

export interface OutboundJournalOptions {
  room: string
  clientId: string
  storageKey?: string
  idFactory: () => string
  storage: OutboundJournalStorage
  transport: OutboundJournalTransport
  scheduler: RetryScheduler
  retry?: Partial<OutboundRetryPolicy>
  hooks?: OutboundJournalHooks
}

export class OutboundJournalOpenError extends Error {
  constructor(readonly code: 'storage' | 'corrupt' | 'version' | 'identity', message: string) {
    super(message)
    this.name = 'OutboundJournalOpenError'
  }
}

export class OutboundJournalCancelledError extends Error {
  constructor(readonly idempotencyKey: string) {
    super(`Outbound journal entry ${idempotencyKey} was cancelled`)
    this.name = 'OutboundJournalCancelledError'
  }
}

interface Waiter { resolve(seq: number): void; reject(error: unknown): void }

const DEFAULT_RETRY: OutboundRetryPolicy = { initialDelayMs: 250, multiplier: 2, maxDelayMs: 30_000 }
const clone = <T>(value: T): T => JSON.parse(JSON.stringify(value)) as T
const isSafeSeq = (value: unknown): value is number => Number.isSafeInteger(value) && (value as number) >= 0

function validOps(value: unknown): value is OpRecord[] {
  return Array.isArray(value) && value.length > 0 && value.every((op) =>
    !!op && typeof op === 'object' && typeof op.id === 'string' && !!op.id.trim()
    && !!op.params && typeof op.params === 'object' && !Array.isArray(op.params),
  )
}

function validateSnapshot(value: unknown): value is OutboundJournalSnapshotV1 {
  if (!value || typeof value !== 'object') return false
  const input = value as Partial<OutboundJournalSnapshotV1>
  if (input.protocol !== OUTBOUND_JOURNAL_PROTOCOL || typeof input.room !== 'string' || !input.room.trim() || typeof input.clientId !== 'string' || !input.clientId.trim()) return false
  if (!isSafeSeq(input.lastServerSeq) || !Number.isSafeInteger(input.nextOrdinal) || input.nextOrdinal! < 1 || !Array.isArray(input.entries)) return false
  let previousOrdinal = 0
  const keys = new Set<string>()
  for (const entry of input.entries) {
    if (!entry || typeof entry.idempotencyKey !== 'string' || !/^[A-Za-z0-9][A-Za-z0-9._:-]{0,191}$/.test(entry.idempotencyKey)) return false
    if (keys.has(entry.idempotencyKey) || !Number.isSafeInteger(entry.ordinal) || entry.ordinal <= previousOrdinal || entry.ordinal >= input.nextOrdinal!) return false
    if (!isSafeSeq(entry.createdBaseSeq) || !Number.isSafeInteger(entry.attempts) || entry.attempts < 0 || !validOps(entry.ops)) return false
    keys.add(entry.idempotencyKey)
    previousOrdinal = entry.ordinal
  }
  return true
}

function parseSnapshot(raw: string): unknown {
  try { return JSON.parse(raw) } catch { throw new OutboundJournalOpenError('corrupt', 'Outbound journal is not valid JSON') }
}

/** Durable, ordered outbound submissions over a host-owned transport. */
export class DurableOutboundJournal {
  private snapshotValue: OutboundJournalSnapshotV1
  private statusValue: OutboundJournalStatus = 'closed'
  private online = false
  private opened = false
  private disposed = false
  private retryAttempt = 0
  private nextRetryMs: number | null = null
  private retryTask?: { cancel(): void }
  private activeAbort?: AbortController
  private pumpPromise?: Promise<void>
  private resyncPromise?: Promise<boolean>
  private mutationTail: Promise<void> = Promise.resolve()
  private submissionRevision = 0
  private readonly waiters = new Map<string, Waiter>()
  /** In-process transaction gate. A restarted journal safely replays every
   * durable entry because local undo history itself is not crash-persisted. */
  private readonly deferred = new Set<string>()
  private readonly retry: OutboundRetryPolicy
  private readonly storageKey: string

  constructor(private readonly options: OutboundJournalOptions) {
    if (!options.room.trim() || !options.clientId.trim()) throw new TypeError('room and clientId are required')
    this.retry = { ...DEFAULT_RETRY, ...options.retry }
    if (!Number.isFinite(this.retry.initialDelayMs) || this.retry.initialDelayMs < 0 || !Number.isFinite(this.retry.multiplier) || this.retry.multiplier < 1 || !Number.isFinite(this.retry.maxDelayMs) || this.retry.maxDelayMs < this.retry.initialDelayMs) {
      throw new TypeError('Invalid outbound retry policy')
    }
    this.storageKey = options.storageKey ?? `injoffice:collab:outbound:${options.room}:${options.clientId}`
    this.snapshotValue = this.emptySnapshot()
  }

  get state(): OutboundJournalState {
    return {
      status: this.statusValue,
      pending: this.snapshotValue.entries.length,
      lastServerSeq: this.snapshotValue.lastServerSeq,
      retryAttempt: this.retryAttempt,
      nextRetryMs: this.nextRetryMs,
    }
  }

  pending(): OutboundJournalEntry[] { return clone(this.snapshotValue.entries) }

  async open(): Promise<void> {
    this.assertNotDisposed()
    if (this.opened) return
    let raw: string | null
    try { raw = await this.options.storage.read(this.storageKey) } catch (error) {
      throw new OutboundJournalOpenError('storage', `Cannot read outbound journal: ${String(error)}`)
    }
    this.assertNotDisposed()
    if (raw === null) {
      try { await this.persist(this.snapshotValue) } catch (error) {
        throw new OutboundJournalOpenError('storage', String(error))
      }
    } else {
      const parsed = parseSnapshot(raw)
      if ((parsed as { protocol?: unknown } | null)?.protocol !== OUTBOUND_JOURNAL_PROTOCOL) {
        this.statusValue = 'blocked'
        throw new OutboundJournalOpenError('version', 'Unsupported outbound journal protocol')
      }
      if (!validateSnapshot(parsed)) {
        this.statusValue = 'blocked'
        throw new OutboundJournalOpenError('corrupt', 'Outbound journal failed validation')
      }
      if (parsed.room !== this.options.room || parsed.clientId !== this.options.clientId) {
        this.statusValue = 'blocked'
        throw new OutboundJournalOpenError('identity', 'Outbound journal belongs to another room or client')
      }
      this.snapshotValue = clone(parsed)
    }
    this.assertNotDisposed()
    this.opened = true
    this.setStatus('offline')
  }

  /** Durably enqueue one batch and return its acknowledgement separately. */
  async enqueue(
    ops: OpRecord[],
    createdBaseSeq = this.snapshotValue.lastServerSeq,
    options: OutboundJournalEnqueueOptions = {},
  ): Promise<OutboundJournalReceipt> {
    if (!validOps(ops) || !isSafeSeq(createdBaseSeq)) throw new TypeError('Invalid outbound submission')
    this.assertUsable()
    this.submissionRevision++
    const receipt = await this.mutate(async () => {
      this.assertUsable()
      const token = this.options.idFactory()
      if (!/^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/.test(token) || this.snapshotValue.nextOrdinal >= Number.MAX_SAFE_INTEGER) {
        throw new TypeError('idFactory returned an invalid token or the journal ordinal is exhausted')
      }
      // Ordinal persistence prevents reuse even if a host token factory repeats.
      // Servers scope idempotency by (room, clientId, idempotencyKey).
      const idempotencyKey = `${this.snapshotValue.nextOrdinal}:${token}`
      const entry: OutboundJournalEntry = {
        idempotencyKey,
        ordinal: this.snapshotValue.nextOrdinal,
        createdBaseSeq,
        attempts: 0,
        ops: clone(ops),
      }
      const next = clone(this.snapshotValue)
      next.nextOrdinal++
      next.entries.push(entry)
      await this.persist(next)
      if (this.disposed) throw new OutboundJournalCancelledError(idempotencyKey)
      this.snapshotValue = next
      if (options.deferSend) this.deferred.add(idempotencyKey)
      this.emit()
      const acknowledged = new Promise<number>((resolve, reject) => this.waiters.set(idempotencyKey, { resolve, reject }))
      return {
        idempotencyKey,
        entry: clone(entry),
        acknowledged,
        release: () => this.release(idempotencyKey),
      }
    })
    if (!options.deferSend && this.online && this.statusValue === 'ready') void this.pump()
    return receipt
  }

  /** Compatible with SubmitQueue's public send callback. */
  async submit(ops: OpRecord[], createdBaseSeq = this.snapshotValue.lastServerSeq): Promise<number> {
    const receipt = await this.enqueue(ops, createdBaseSeq)
    return receipt.acknowledged
  }

  async connect(): Promise<boolean> {
    this.assertOpen()
    if (this.online && (this.statusValue === 'ready' || this.statusValue === 'sending')) return true
    this.online = true
    this.cancelRetry()
    if (this.resyncPromise) await this.resyncPromise
    if (!this.online || this.disposed) return false
    this.cancelRetry()
    return this.resynchronize()
  }

  disconnect(): void {
    if (!this.opened || this.disposed) return
    this.online = false
    this.activeAbort?.abort()
    this.activeAbort = undefined
    this.cancelRetry()
    this.setStatus('offline')
  }

  async cancel(idempotencyKey: string): Promise<boolean> {
    this.assertOpen()
    if (this.snapshotValue.entries[0]?.idempotencyKey === idempotencyKey) this.activeAbort?.abort()
    const { removed, stillBlocked } = await this.mutate(async () => {
      this.assertOpen()
      const wasBlocked = this.statusValue === 'blocked'
      const index = this.snapshotValue.entries.findIndex((entry) => entry.idempotencyKey === idempotencyKey)
      if (index < 0) return { removed: false, stillBlocked: wasBlocked }
      const next = clone(this.snapshotValue)
      next.entries.splice(index, 1)
      await this.persist(next)
      if (this.disposed) return { removed: false, stillBlocked: false }
      this.snapshotValue = next
      this.deferred.delete(idempotencyKey)
      this.waiters.get(idempotencyKey)?.reject(new OutboundJournalCancelledError(idempotencyKey))
      this.waiters.delete(idempotencyKey)
      return { removed: true, stillBlocked: wasBlocked && next.entries.length > 0 }
    })
    if (!removed) return false
    if (this.disposed) return true
    this.setStatus(stillBlocked ? 'blocked' : this.online ? 'ready' : 'offline')
    if (this.online && !stillBlocked) void this.pump()
    return true
  }

  dispose(): void {
    if (this.disposed) return
    this.disposed = true
    this.online = false
    this.activeAbort?.abort()
    this.cancelRetry()
    for (const [key, waiter] of this.waiters) waiter.reject(new OutboundJournalCancelledError(key))
    this.waiters.clear()
    this.deferred.clear()
    this.setStatus('disposed')
  }

  private resynchronize(): Promise<boolean> {
    if (this.resyncPromise) return this.resyncPromise
    this.resyncPromise = this.runResynchronize().finally(() => { this.resyncPromise = undefined })
    return this.resyncPromise
  }

  private async runResynchronize(): Promise<boolean> {
    if (!this.online || this.disposed) return false
    this.setStatus('resyncing')
    await this.mutate(async () => undefined)
    if (!this.online || this.disposed) return false
    const abort = new AbortController()
    this.activeAbort = abort
    let result: JournalResyncResult
    const requestPending = this.pending()
    const requestAfterSeq = this.snapshotValue.lastServerSeq
    const requestSubmissionRevision = this.submissionRevision
    try {
      result = await this.options.transport.resync({
        room: this.options.room,
        clientId: this.options.clientId,
        afterSeq: requestAfterSeq,
        pending: requestPending,
      }, abort.signal)
    } catch (error) {
      if (this.activeAbort === abort) this.activeAbort = undefined
      if (!abort.signal.aborted && this.online) this.scheduleRetry(error)
      return false
    }
    if (this.activeAbort === abort) this.activeAbort = undefined
    if (this.disposed || !this.online || abort.signal.aborted) return false
    if (!this.validResync(result, requestPending, requestAfterSeq)) {
      this.setStatus('blocked')
      this.options.hooks?.onBlocked?.('invalid-resync', result)
      return false
    }
    if (result.replay === 'blocked' && this.snapshotValue.entries.length > 0) {
      this.setStatus('blocked')
      this.options.hooks?.onBlocked?.('unsafe-replay', result)
      return false
    }
    let remaining = 0
    let submittedDuringResync = false
    try { await this.mutate(async () => {
      const next = clone(this.snapshotValue)
      next.lastServerSeq = Math.max(next.lastServerSeq, result.headSeq)
      if (result.rebased) {
        const byKey = new Map(result.rebased.map((entry) => [entry.idempotencyKey, entry.ops]))
        for (const entry of next.entries) {
          const ops = byKey.get(entry.idempotencyKey)
          if (ops) entry.ops = clone(ops)
        }
      }
      await this.persist(next)
      this.snapshotValue = next
      remaining = next.entries.length
      submittedDuringResync = this.submissionRevision !== requestSubmissionRevision
    }) } catch (error) {
      this.scheduleRetry(error)
      return false
    }
    if (submittedDuringResync) {
      this.scheduleRetry('submission arrived during resync')
      return false
    }
    if (remaining === 0) this.retryAttempt = 0
    this.nextRetryMs = null
    this.setStatus('ready')
    void this.pump()
    return true
  }

  private async pump(): Promise<void> {
    if (this.pumpPromise) return this.pumpPromise
    this.pumpPromise = this.runPump().finally(() => {
      this.pumpPromise = undefined
      const head = this.snapshotValue.entries[0]
      if (this.online && !this.disposed && this.statusValue === 'ready' && head && !this.deferred.has(head.idempotencyKey)) void this.pump()
    })
    return this.pumpPromise
  }

  private async runPump(): Promise<void> {
    while (this.online && !this.disposed && this.statusValue === 'ready' && this.snapshotValue.entries.length > 0) {
      if (this.deferred.has(this.snapshotValue.entries[0].idempotencyKey)) return
      let entry: OutboundJournalEntry | null = null
      try { entry = await this.mutate(async () => {
        if (!this.online || this.disposed || this.statusValue !== 'ready' || this.snapshotValue.entries.length === 0) return null
        const attempted = clone(this.snapshotValue)
        attempted.entries[0].attempts++
        await this.persist(attempted)
        if (this.disposed || !this.online) return null
        this.snapshotValue = attempted
        return clone(attempted.entries[0])
      }) } catch (error) { this.scheduleRetry(error); return }
      if (!entry) return
      this.setStatus('sending')
      const abort = new AbortController()
      this.activeAbort = abort
      let ack: JournalSubmitAck
      try {
        ack = await this.options.transport.submit({
          room: this.options.room,
          clientId: this.options.clientId,
          idempotencyKey: entry.idempotencyKey,
          baseSeq: this.snapshotValue.lastServerSeq,
          ops: clone(entry.ops),
        }, abort.signal)
      } catch (error) {
        if (this.activeAbort === abort) this.activeAbort = undefined
        if (!abort.signal.aborted && this.online) this.scheduleRetry(error)
        else if (!this.online) this.setStatus('offline')
        return
      }
      if (this.activeAbort === abort) this.activeAbort = undefined
      if (this.disposed || !this.online || abort.signal.aborted) return
      if (this.snapshotValue.entries[0]?.idempotencyKey !== entry.idempotencyKey) {
        this.setStatus('ready')
        continue
      }
      if (!ack || ack.idempotencyKey !== entry.idempotencyKey || !isSafeSeq(ack.seq) || ack.seq < 1) {
        this.setStatus('blocked')
        this.options.hooks?.onBlocked?.('invalid-resync', ack)
        return
      }
      if (ack.seq <= this.snapshotValue.lastServerSeq && ack.deduplicated !== true) {
        this.setStatus('blocked')
        this.options.hooks?.onBlocked?.('invalid-resync', ack)
        return
      }
      let committed = false
      try { committed = await this.mutate(async () => {
        if (this.snapshotValue.entries[0]?.idempotencyKey !== entry.idempotencyKey) return false
        const next = clone(this.snapshotValue)
        next.entries.shift()
        next.lastServerSeq = Math.max(next.lastServerSeq, ack.seq)
        await this.persist(next)
        this.snapshotValue = next
        return true
      }) } catch (error) {
        this.scheduleRetry(error)
        return
      }
      if (!committed) { this.setStatus('ready'); continue }
      this.retryAttempt = 0
      this.nextRetryMs = null
      this.waiters.get(entry.idempotencyKey)?.resolve(ack.seq)
      this.waiters.delete(entry.idempotencyKey)
      this.options.hooks?.onAck?.(clone(entry), { ...ack })
      this.setStatus('ready')
    }
  }

  private validResync(result: JournalResyncResult, pending: readonly OutboundJournalEntry[], afterSeq: number): boolean {
    if (!result || !isSafeSeq(result.headSeq) || result.headSeq < afterSeq || (result.replay !== 'safe' && result.replay !== 'blocked')) return false
    if (!result.rebased) return true
    if (result.rebased.length !== pending.length) return false
    const expected = pending.map((entry) => entry.idempotencyKey)
    return result.rebased.every((entry, index) => entry?.idempotencyKey === expected[index] && validOps(entry.ops))
  }

  private scheduleRetry(_error: unknown): void {
    if (!this.online || this.disposed || this.retryTask) return
    this.retryAttempt++
    const delay = Math.min(this.retry.maxDelayMs, this.retry.initialDelayMs * this.retry.multiplier ** (this.retryAttempt - 1))
    this.nextRetryMs = delay
    this.setStatus('retry-wait')
    this.retryTask = this.options.scheduler.schedule(delay, () => {
      this.retryTask = undefined
      this.nextRetryMs = null
      if (this.online && !this.disposed) void this.resynchronize()
    })
  }

  private cancelRetry(): void {
    this.retryTask?.cancel()
    this.retryTask = undefined
    this.nextRetryMs = null
  }

  private async persist(snapshot: OutboundJournalSnapshotV1): Promise<void> {
    try { await this.options.storage.write(this.storageKey, JSON.stringify(snapshot)) } catch (error) {
      throw new Error(`Cannot persist outbound journal: ${String(error)}`)
    }
  }

  private mutate<T>(operation: () => Promise<T>): Promise<T> {
    const result = this.mutationTail.then(operation, operation)
    this.mutationTail = result.then(() => undefined, () => undefined)
    return result
  }

  private emptySnapshot(): OutboundJournalSnapshotV1 {
    return {
      protocol: OUTBOUND_JOURNAL_PROTOCOL,
      room: this.options.room,
      clientId: this.options.clientId,
      lastServerSeq: 0,
      nextOrdinal: 1,
      entries: [],
    }
  }

  private assertNotDisposed(): void { if (this.disposed) throw new Error('Outbound journal is disposed') }
  private assertOpen(): void { this.assertNotDisposed(); if (!this.opened) throw new Error('Outbound journal is not open') }
  private assertUsable(): void { this.assertOpen(); if (this.statusValue === 'blocked') throw new Error('Outbound journal is blocked') }
  private setStatus(status: OutboundJournalStatus): void {
    if (this.disposed && status !== 'disposed') return
    this.statusValue = status
    this.emit()
  }
  private release(idempotencyKey: string): void {
    this.deferred.delete(idempotencyKey)
    if (this.online && this.statusValue === 'ready') void this.pump()
  }
  private emit(): void { this.options.hooks?.onState?.(this.state) }
}

/** Public-boundary adapter for `new SubmitQueue(send, hooks)`. */
export function createJournalSubmitter(journal: DurableOutboundJournal, baseSeq: () => number): (ops: OpRecord[]) => Promise<number> {
  return (ops) => journal.submit(ops, baseSeq())
}
