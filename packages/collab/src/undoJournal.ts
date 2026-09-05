import { DurableOutboundJournal, type OutboundJournalReceipt } from './offlineJournal'
import type { OpRecord } from './sync'
import {
  CollaborativeUndoManager,
  type CollaborativeUndoEntry,
  type UndoBlockReason,
  type UndoRebaseReport,
} from './undo'
import type { SheetSelection } from './types'

export interface CollaborativeUndoRemoteEntry {
  sequence: number
  ops: OpRecord[]
}

export type CollaborativeUndoExecutionOrigin = 'remote' | 'undo' | 'redo'

export interface CollaborativeUndoExecutionContext {
  origin: CollaborativeUndoExecutionOrigin
  entryId?: string
  sequence?: number
  /** Prevent a host command listener from journaling this batch a second time. */
  fromCollaborativeUndo: boolean
  fromCollab: boolean
}

export interface CollaborativeUndoJournalHost {
  /** Apply the complete batch atomically. False means no mutation was applied. */
  execute(ops: readonly OpRecord[], context: CollaborativeUndoExecutionContext): boolean | Promise<boolean>
  restoreSelection?(selection: SheetSelection): void
}

export type CollaborativeUndoJournalBlockCode =
  | 'sequence-gap'
  | 'execution-failed'
  | 'history-failed'
  | 'journal-failed'
  | 'stale-plan'
  | 'acknowledgement-gap'

export interface CollaborativeUndoJournalState {
  sequence: number
  pendingHistoryEntries: number
  status: 'active' | 'blocked' | 'disposed'
  blockCode?: CollaborativeUndoJournalBlockCode
}

export interface CollaborativeUndoJournalOptions {
  initialSequence?: number
  onState?: (state: CollaborativeUndoJournalState) => void
  onBlocked?: (code: CollaborativeUndoJournalBlockCode, detail?: unknown) => void
  onAcknowledged?: (entryId: string, sequence: number) => void
}

export interface CollaborativeUndoJournalReceipt {
  entryId: string
  idempotencyKey: string
  acknowledged: Promise<number>
}

export type CollaborativeUndoJournalPlanResult =
  | { status: 'empty'; direction: 'undo' | 'redo' }
  | { status: 'blocked'; direction: 'undo' | 'redo'; entryId?: string; reasons?: readonly UndoBlockReason[]; code?: CollaborativeUndoJournalBlockCode }
  | { status: 'neutralized'; direction: 'undo' | 'redo'; entryId: string }
  | ({ status: 'applied'; direction: 'undo' | 'redo' } & CollaborativeUndoJournalReceipt)

export type CollaborativeUndoRemoteResult =
  | { status: 'applied'; sequence: number; report: UndoRebaseReport }
  | { status: 'duplicate'; sequence: number }
  | { status: 'blocked'; sequence: number; code: CollaborativeUndoJournalBlockCode }

const clone = <T>(value: T): T => JSON.parse(JSON.stringify(value)) as T

function validSequence(value: unknown): value is number {
  return Number.isSafeInteger(value) && (value as number) >= 0
}

function validOps(value: unknown): value is OpRecord[] {
  if (!Array.isArray(value) || value.length === 0) return false
  try {
    JSON.stringify(value)
  } catch {
    return false
  }
  return value.every((op) => !!op && typeof op === 'object' && typeof op.id === 'string' && !!op.id.trim()
    && !!op.params && typeof op.params === 'object' && !Array.isArray(op.params))
}

/**
 * Connects the causal undo planner to the durable outbound journal.
 *
 * Local inverses are durably enqueued but held from transport until the host
 * atomically executes them and the revision-checked plan commits. Ordered
 * remote batches execute through the same host and rebase history with a
 * distinction between acknowledged local actions and still-outbound actions.
 */
export class CollaborativeUndoJournalCoordinator {
  private sequenceValue: number
  private statusValue: 'active' | 'blocked' | 'disposed' = 'active'
  private blockCodeValue: CollaborativeUndoJournalBlockCode | undefined
  private readonly pendingCounts = new Map<string, number>()
  private tail: Promise<void> = Promise.resolve()

  constructor(
    readonly undo: CollaborativeUndoManager,
    readonly journal: DurableOutboundJournal,
    private readonly host: CollaborativeUndoJournalHost,
    private readonly options: CollaborativeUndoJournalOptions = {},
  ) {
    const initial = options.initialSequence ?? journal.state.lastServerSeq
    if (!validSequence(initial)) throw new TypeError('initialSequence must be a non-negative safe integer')
    this.sequenceValue = initial
  }

  get state(): CollaborativeUndoJournalState {
    return {
      sequence: this.sequenceValue,
      pendingHistoryEntries: this.pendingCounts.size,
      status: this.statusValue,
      ...(this.blockCodeValue ? { blockCode: this.blockCodeValue } : {}),
    }
  }

  /** Record an already-applied local action only after its forward batch is
   * durably accepted by the outbound journal. */
  recordLocal(entry: CollaborativeUndoEntry, createdBaseSeq = this.sequenceValue): Promise<CollaborativeUndoJournalReceipt> {
    return this.serialize(async () => {
      this.assertActive()
      if (!validOps(entry.redoOps) || !validSequence(createdBaseSeq)) throw new TypeError('A local action requires a valid non-empty forward batch and base sequence')
      let outbound: OutboundJournalReceipt
      try {
        outbound = await this.journal.enqueue(entry.redoOps, createdBaseSeq, { deferSend: true })
      } catch (error) {
        this.setBlocked('journal-failed', error)
        throw error
      }
      let entryId: string
      try {
        entryId = this.undo.record(entry)
      } catch (error) {
        await this.cancel(outbound)
        this.setBlocked('history-failed', error)
        throw error
      }
      this.track(entryId, outbound)
      outbound.release()
      this.emit()
      return { entryId, idempotencyKey: outbound.idempotencyKey, acknowledged: outbound.acknowledged }
    })
  }

  undoOnce(createdBaseSeq = this.sequenceValue): Promise<CollaborativeUndoJournalPlanResult> {
    return this.executePlan('undo', createdBaseSeq)
  }

  redoOnce(createdBaseSeq = this.sequenceValue): Promise<CollaborativeUndoJournalPlanResult> {
    return this.executePlan('redo', createdBaseSeq)
  }

  receiveRemote(entry: CollaborativeUndoRemoteEntry): Promise<CollaborativeUndoRemoteResult> {
    return this.serialize(async () => {
      if (this.statusValue !== 'active') return { status: 'blocked', sequence: this.sequenceValue, code: this.blockCodeValue ?? 'execution-failed' }
      if (!entry || !validSequence(entry.sequence) || entry.sequence < 1 || !validOps(entry.ops)) {
        return this.block('sequence-gap', entry)
      }
      if (entry.sequence <= this.sequenceValue) return { status: 'duplicate', sequence: this.sequenceValue }
      if (entry.sequence !== this.sequenceValue + 1) return this.block('sequence-gap', entry)
      const ops = clone(entry.ops)
      let executed = false
      try {
        executed = await this.host.execute(ops, {
          origin: 'remote', sequence: entry.sequence, fromCollaborativeUndo: false, fromCollab: true,
        }) === true
      } catch (error) {
        return this.block('execution-failed', error)
      }
      if (!executed) return this.block('execution-failed', entry)
      const report = this.undo.rebaseRemote(ops, { remoteBeforeEntryIds: [...this.pendingCounts.keys()] })
      this.sequenceValue = entry.sequence
      this.emit()
      return { status: 'applied', sequence: this.sequenceValue, report }
    })
  }

  /** Recover only after the host has loaded authoritative state and resolved
   * every durable pending entry. Local history is cleared because a hard
   * snapshot reload cannot prove its inverse preconditions still hold. */
  resetAfterResync(sequence: number): boolean {
    if (this.statusValue === 'disposed' || !validSequence(sequence) || this.pendingCounts.size > 0 || this.journal.pending().length > 0) return false
    this.undo.clear()
    this.sequenceValue = sequence
    this.statusValue = 'active'
    this.blockCodeValue = undefined
    this.emit()
    return true
  }

  dispose(): void {
    if (this.statusValue === 'disposed') return
    this.statusValue = 'disposed'
    this.blockCodeValue = undefined
    this.pendingCounts.clear()
    this.emit()
  }

  private executePlan(direction: 'undo' | 'redo', createdBaseSeq: number): Promise<CollaborativeUndoJournalPlanResult> {
    return this.serialize(async () => {
      if (this.statusValue !== 'active') return { status: 'blocked', direction, code: this.blockCodeValue ?? 'execution-failed' }
      if (!validSequence(createdBaseSeq)) throw new TypeError('createdBaseSeq must be a non-negative safe integer')
      const plan = direction === 'undo' ? this.undo.planUndo() : this.undo.planRedo()
      if (plan.status === 'empty') return plan
      if (plan.status === 'blocked') return { status: 'blocked', direction, entryId: plan.entryId, reasons: plan.reasons }
      if (plan.status === 'neutralized') {
        if (!this.undo.commit(plan)) return this.blockPlan(direction, plan.entryId, 'stale-plan')
        this.restoreSelection(plan.selection)
        this.emit()
        return { status: 'neutralized', direction, entryId: plan.entryId }
      }

      let outbound: OutboundJournalReceipt
      try {
        outbound = await this.journal.enqueue([...plan.ops], createdBaseSeq, { deferSend: true })
      } catch (error) {
        return this.blockPlan(direction, plan.entryId, 'journal-failed', error)
      }
      let executed = false
      try {
        executed = await this.host.execute(clone(plan.ops), {
          origin: direction,
          entryId: plan.entryId,
          fromCollaborativeUndo: true,
          fromCollab: false,
        }) === true
      } catch (error) {
        await this.cancel(outbound)
        return this.blockPlan(direction, plan.entryId, 'execution-failed', error)
      }
      if (!executed) {
        await this.cancel(outbound)
        return this.blockPlan(direction, plan.entryId, 'execution-failed')
      }
      if (!this.undo.commit(plan)) {
        await this.cancel(outbound)
        return this.blockPlan(direction, plan.entryId, 'stale-plan')
      }
      this.restoreSelection(plan.selection)
      this.track(plan.entryId, outbound)
      outbound.release()
      this.emit()
      return {
        status: 'applied', direction, entryId: plan.entryId,
        idempotencyKey: outbound.idempotencyKey, acknowledged: outbound.acknowledged,
      }
    })
  }

  private track(entryId: string, receipt: OutboundJournalReceipt): void {
    this.pendingCounts.set(entryId, (this.pendingCounts.get(entryId) ?? 0) + 1)
    void receipt.acknowledged.then(
      (sequence) => this.serialize(async () => { if (this.statusValue !== 'disposed') this.acknowledge(entryId, sequence) }),
      (error) => this.serialize(async () => { if (this.statusValue !== 'disposed') this.block('journal-failed', error) }),
    )
  }

  private acknowledge(entryId: string, sequence: number): void {
    const count = this.pendingCounts.get(entryId) ?? 0
    if (count <= 1) this.pendingCounts.delete(entryId)
    else this.pendingCounts.set(entryId, count - 1)
    if (!validSequence(sequence) || sequence > this.sequenceValue + 1) {
      this.block('acknowledgement-gap', { entryId, sequence })
      return
    }
    if (sequence === this.sequenceValue + 1) this.sequenceValue = sequence
    try { this.options.onAcknowledged?.(entryId, sequence) } catch { /* observer isolation */ }
    this.emit()
  }

  private async cancel(receipt: OutboundJournalReceipt): Promise<void> {
    // Attach a rejection observer before cancel() rejects the journal waiter.
    void receipt.acknowledged.catch(() => undefined)
    try {
      if (!await this.journal.cancel(receipt.idempotencyKey)) this.block('journal-failed', receipt.idempotencyKey)
    } catch (error) {
      this.block('journal-failed', error)
    }
  }

  private restoreSelection(selection: SheetSelection | null): void {
    if (!selection) return
    try { this.host.restoreSelection?.(clone(selection)) } catch { /* selection is advisory */ }
  }

  private blockPlan(
    direction: 'undo' | 'redo',
    entryId: string,
    code: CollaborativeUndoJournalBlockCode,
    detail?: unknown,
  ): CollaborativeUndoJournalPlanResult {
    this.setBlocked(code, detail)
    return { status: 'blocked', direction, entryId, code }
  }

  private block(code: CollaborativeUndoJournalBlockCode, detail?: unknown): { status: 'blocked'; sequence: number; code: CollaborativeUndoJournalBlockCode } {
    this.setBlocked(code, detail)
    return { status: 'blocked', sequence: this.sequenceValue, code }
  }

  private setBlocked(code: CollaborativeUndoJournalBlockCode, detail?: unknown): void {
    if (this.statusValue === 'disposed') return
    this.statusValue = 'blocked'
    this.blockCodeValue = code
    try { this.options.onBlocked?.(code, detail) } catch { /* observer isolation */ }
    this.emit()
  }

  private assertActive(): void {
    if (this.statusValue !== 'active') throw new Error(`Collaborative undo journal is ${this.statusValue}`)
  }

  private serialize<T>(operation: () => Promise<T>): Promise<T> {
    const result = this.tail.then(operation, operation)
    this.tail = result.then(() => undefined, () => undefined)
    return result
  }

  private emit(): void {
    try { this.options.onState?.(this.state) } catch { /* observer isolation */ }
  }
}
