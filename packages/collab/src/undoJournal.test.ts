import { describe, expect, it, vi } from 'vitest'
import {
  DurableOutboundJournal,
  type JournalSubmitRequest,
  type OutboundJournalStorage,
  type OutboundJournalTransport,
  type RetryScheduler,
} from './offlineJournal'
import { SET_RANGE_VALUES, type OpRecord } from './sync'
import { CollaborativeUndoManager } from './undo'
import { CollaborativeUndoJournalCoordinator, type CollaborativeUndoExecutionContext } from './undoJournal'

const range = (startRow: number, endRow = startRow) => ({ startRow, endRow, startColumn: 0, endColumn: 0 })
const cells = (row: number, value: string): OpRecord => ({
  id: SET_RANGE_VALUES,
  params: { unitId: 'wb', subUnitId: 's1', cellValue: { [row]: { 0: { v: value } } } },
})
const insertRows = (start: number, count: number): OpRecord => ({
  id: 'sheet.mutation.insert-row',
  params: { unitId: 'wb', subUnitId: 's1', range: range(start, start + count - 1) },
})
const drain = async (turns = 30): Promise<void> => { for (let index = 0; index < turns; index++) await Promise.resolve() }

class MemoryStorage implements OutboundJournalStorage {
  value: string | null = null
  read(): string | null { return this.value }
  write(_key: string, value: string): void { this.value = value }
}

const scheduler: RetryScheduler = { schedule: () => ({ cancel: () => undefined }) }

function setup(options: {
  transport?: OutboundJournalTransport
  execute?: (ops: readonly OpRecord[], context: CollaborativeUndoExecutionContext) => boolean | Promise<boolean>
  initialSequence?: number
} = {}) {
  let token = 0
  let serverSequence = 0
  const sent: JournalSubmitRequest[] = []
  const transport: OutboundJournalTransport = options.transport ?? {
    resync: async () => ({ headSeq: serverSequence, replay: 'safe' }),
    submit: async (request) => {
      sent.push(request)
      return { idempotencyKey: request.idempotencyKey, seq: ++serverSequence }
    },
  }
  const journal = new DurableOutboundJournal({
    room: 'book', clientId: 'alice', storage: new MemoryStorage(), scheduler, transport,
    idFactory: () => `token-${++token}`,
  })
  const undo = new CollaborativeUndoManager()
  const executions: Array<{ ops: readonly OpRecord[]; context: CollaborativeUndoExecutionContext }> = []
  const coordinator = new CollaborativeUndoJournalCoordinator(undo, journal, {
    execute: async (ops, context) => {
      executions.push({ ops, context })
      return options.execute ? options.execute(ops, context) : true
    },
  }, { initialSequence: options.initialSequence })
  return { coordinator, executions, journal, sent, undo }
}

describe('CollaborativeUndoJournalCoordinator', () => {
  it('rebases pending offline history through earlier remote structure without losing local ownership', async () => {
    let serverSequence = 1
    const sent: JournalSubmitRequest[] = []
    const transport: OutboundJournalTransport = {
      resync: async (request) => ({
        headSeq: 1,
        replay: 'safe',
        rebased: request.pending.map((entry) => ({ idempotencyKey: entry.idempotencyKey, ops: [cells(5, 'local')] })),
      }),
      submit: async (request) => {
        sent.push(request)
        return { idempotencyKey: request.idempotencyKey, seq: ++serverSequence }
      },
    }
    const subject = setup({
      transport,
      execute: (_ops, context) => {
        if (context.origin === 'undo') expect(sent).toHaveLength(1)
        return true
      },
    })
    await subject.journal.open()
    const local = await subject.coordinator.recordLocal({
      id: 'offline-edit',
      undoOps: [cells(3, 'before')],
      redoOps: [cells(3, 'local')],
      selectionBefore: { sheet: 's1', ranges: [[3, 0, 3, 0]], active: [3, 0] },
    })

    const remote = await subject.coordinator.receiveRemote({
      sequence: 1,
      ops: [cells(3, 'remote-earlier'), insertRows(0, 2)],
    })
    expect(remote).toMatchObject({ status: 'applied', sequence: 1, report: { rebased: 1, neutralized: 0, blocked: 0 } })
    expect(subject.undo.planUndo()).toMatchObject({
      status: 'ready',
      ops: [{ params: { cellValue: { 5: { 0: { v: 'before' } } } } }],
      selection: { ranges: [[5, 0, 5, 0]], active: [5, 0] },
    })

    await subject.journal.connect()
    expect(await local.acknowledged).toBe(2)
    await drain()
    expect(sent[0]).toMatchObject({ baseSeq: 1, ops: [cells(5, 'local')] })
    expect(subject.coordinator.state).toMatchObject({ sequence: 2, pendingHistoryEntries: 0, status: 'active' })

    const undone = await subject.coordinator.undoOnce()
    expect(undone).toMatchObject({ status: 'applied', direction: 'undo', entryId: 'offline-edit' })
    if (undone.status !== 'applied') throw new Error('expected applied undo')
    expect(await undone.acknowledged).toBe(3)
    await drain()
    expect(sent[1]).toMatchObject({ baseSeq: 2, ops: [cells(5, 'before')] })
    expect(subject.undo).toMatchObject({ undoDepth: 0, redoDepth: 1 })

    const redone = await subject.coordinator.redoOnce()
    expect(redone).toMatchObject({ status: 'applied', direction: 'redo', entryId: 'offline-edit' })
    if (redone.status !== 'applied') throw new Error('expected applied redo')
    expect(await redone.acknowledged).toBe(4)
    await drain()
    expect(sent[2]).toMatchObject({ baseSeq: 3, ops: [cells(5, 'local')] })
  })

  it('neutralizes acknowledged history after a later remote cell takes ownership', async () => {
    const subject = setup()
    await subject.journal.open()
    await subject.journal.connect()
    const local = await subject.coordinator.recordLocal({ id: 'edit', undoOps: [cells(1, 'before')], redoOps: [cells(1, 'local')] })
    expect(await local.acknowledged).toBe(1)
    await drain()

    expect(await subject.coordinator.receiveRemote({ sequence: 2, ops: [cells(1, 'remote-later')] })).toMatchObject({
      status: 'applied', report: { neutralized: 1 },
    })
    expect(await subject.coordinator.undoOnce()).toEqual({ status: 'neutralized', direction: 'undo', entryId: 'edit' })
    expect(subject.sent).toHaveLength(1)
    expect(subject.executions.map(({ context }) => context.origin)).toEqual(['remote'])
  })

  it('holds undo transport until atomic host execution succeeds and cancels on refusal', async () => {
    const subject = setup({ execute: (_ops, context) => context.origin !== 'undo' })
    await subject.journal.open()
    const local = await subject.coordinator.recordLocal({ id: 'edit', undoOps: [cells(2, 'before')], redoOps: [cells(2, 'local')] })
    void local.acknowledged.catch(() => undefined)
    expect(subject.journal.pending()).toHaveLength(1)

    expect(await subject.coordinator.undoOnce()).toMatchObject({
      status: 'blocked', direction: 'undo', entryId: 'edit', code: 'execution-failed',
    })
    expect(subject.undo).toMatchObject({ undoDepth: 1, redoDepth: 0 })
    expect(subject.journal.pending()).toMatchObject([{ ops: [cells(2, 'local')] }])
    expect(subject.coordinator.state).toMatchObject({ status: 'blocked', blockCode: 'execution-failed' })
  })

  it('blocks sequence gaps and only resets after pending durable work is resolved', async () => {
    const blocked = vi.fn()
    const subject = setup()
    await subject.journal.open()
    const coordinator = new CollaborativeUndoJournalCoordinator(subject.undo, subject.journal, {
      execute: () => true,
    }, { onBlocked: blocked })
    const local = await coordinator.recordLocal({ id: 'edit', undoOps: [cells(0, 'before')], redoOps: [cells(0, 'local')] })
    void local.acknowledged.catch(() => undefined)

    expect(await coordinator.receiveRemote({ sequence: 2, ops: [cells(3, 'gap')] })).toEqual({
      status: 'blocked', sequence: 0, code: 'sequence-gap',
    })
    expect(coordinator.resetAfterResync(2)).toBe(false)
    await subject.journal.cancel(local.idempotencyKey)
    coordinator.dispose()
    expect(coordinator.resetAfterResync(2)).toBe(false)
    expect(blocked).toHaveBeenCalledWith('sequence-gap', expect.anything())

    const clean = setup()
    await clean.journal.open()
    clean.undo.record({ id: 'discarded-on-hard-resync', undoOps: [cells(1, 'old')], redoOps: [cells(1, 'new')] })
    expect(await clean.coordinator.receiveRemote({ sequence: 3, ops: [cells(2, 'gap')] })).toMatchObject({ status: 'blocked', code: 'sequence-gap' })
    expect(clean.coordinator.resetAfterResync(3)).toBe(true)
    expect(clean.coordinator.state).toEqual({ sequence: 3, pendingHistoryEntries: 0, status: 'active' })
    expect(clean.undo.planUndo()).toEqual({ status: 'empty', direction: 'undo' })
  })

  it('fails closed when an acknowledgement skips unseen ordered entries', async () => {
    const subject = setup({
      transport: {
        resync: async () => ({ headSeq: 0, replay: 'safe' }),
        submit: async (request) => ({ idempotencyKey: request.idempotencyKey, seq: 3 }),
      },
    })
    await subject.journal.open()
    await subject.journal.connect()
    const local = await subject.coordinator.recordLocal({ id: 'edit', undoOps: [cells(1, 'old')], redoOps: [cells(1, 'new')] })
    expect(await local.acknowledged).toBe(3)
    await drain()
    expect(subject.coordinator.state).toMatchObject({ status: 'blocked', blockCode: 'acknowledgement-gap', sequence: 0, pendingHistoryEntries: 0 })
  })

  it('returns planner block reasons without executing or journaling guessed object inverses', async () => {
    const subject = setup()
    await subject.journal.open()
    const local = await subject.coordinator.recordLocal({
      id: 'drawing',
      undoOps: [{ id: 'sheet.mutation.set-drawing', params: { subUnitId: 's1', drawingId: 'd1' } }],
      redoOps: [{ id: 'sheet.mutation.set-drawing', params: { subUnitId: 's1', drawingId: 'd1' } }],
    })
    void local.acknowledged.catch(() => undefined)
    const before = subject.journal.pending().length
    expect(await subject.coordinator.undoOnce()).toMatchObject({
      status: 'blocked',
      reasons: [{ code: 'unsupported-local-operation', opId: 'sheet.mutation.set-drawing' }],
    })
    expect(subject.journal.pending()).toHaveLength(before)
    expect(subject.executions).toEqual([])
  })
})
