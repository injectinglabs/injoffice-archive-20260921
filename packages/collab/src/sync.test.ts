import { describe, expect, it } from 'vitest'
import { MOVE_RANGE, REORDER_RANGE, SET_RANGE_VALUES, SubmitQueue, cellsOf, cloneParams, pruneOwned, rebaseUnit, reconcileRemote, shouldShipMutation, type OpEntry, type OpRecord } from './sync'

const srv = (cells: Record<string, Record<string, unknown>>, sub = 's1') => ({
  id: SET_RANGE_VALUES,
  params: { unitId: 'wb', subUnitId: sub, cellValue: cells },
})

describe('shouldShipMutation', () => {
  const base = { id: SET_RANGE_VALUES, type: 2, params: { unitId: 'wb' } }
  it('ships plain local sheet mutations for our unit', () => {
    expect(shouldShipMutation(base, undefined, 'wb')).toBe(true)
    expect(shouldShipMutation(base, {}, 'wb')).toBe(true)
  })
  it('never echoes remote, snapshot, local-only or sync-only executions', () => {
    expect(shouldShipMutation(base, { fromCollab: true }, 'wb')).toBe(false)
    expect(shouldShipMutation(base, { fromChangeset: true }, 'wb')).toBe(false)
    expect(shouldShipMutation(base, { onlyLocal: true }, 'wb')).toBe(false)
    expect(shouldShipMutation(base, { syncOnly: true }, 'wb')).toBe(false)
  })
  it('skips commands/operations, other units, doc/formula internals and freeze', () => {
    expect(shouldShipMutation({ ...base, type: 0 }, {}, 'wb')).toBe(false)
    expect(shouldShipMutation({ ...base, params: { unitId: 'other' } }, {}, 'wb')).toBe(false)
    expect(shouldShipMutation({ id: 'doc.mutation.rich-text-editing', type: 2, params: { unitId: 'wb' } }, {}, 'wb')).toBe(false)
    expect(shouldShipMutation({ id: 'formula.mutation.set-formula-calculation-result', type: 2, params: {} }, {}, 'wb')).toBe(false)
    expect(shouldShipMutation({ id: 'sheet.mutation.set-frozen', type: 2, params: { unitId: 'wb' } }, {}, 'wb')).toBe(false)
    expect(shouldShipMutation({ id: SET_RANGE_VALUES, type: 2 }, {}, 'wb')).toBe(false)
  })
  it('accepts mutations without a unitId (workbook-wide) and clones params deeply', () => {
    expect(shouldShipMutation({ id: 'sheet.mutation.set-worksheet-name', type: 2, params: { name: 'x' } }, {}, 'wb')).toBe(true)
    const src = { a: { b: 1 } }
    const c = cloneParams(src)
    ;(c.a as { b: number }).b = 2
    expect(src.a.b).toBe(1)
  })
})

describe('rebaseUnit', () => {
  it('rewrites the sender unit id onto ours and leaves unit-less ops alone', () => {
    const ops: OpRecord[] = [srv({ '1': { '1': { v: 1 } } }), { id: 'sheet.mutation.set-worksheet-name', params: { name: 'n' } }]
    const out = rebaseUnit(ops, 'mine')
    expect(out[0].params.unitId).toBe('mine')
    expect(out[0].params.subUnitId).toBe('s1')
    expect(out[1]).toBe(ops[1])
    expect(ops[0].params.unitId).toBe('wb') // input untouched
  })
  it('makes reconcile keys line up across tabs', () => {
    const owned = [{ seq: null, cells: cellsOf({ ...srv({ '1': { '1': { v: 'mine' } } }), params: { ...srv({}).params, unitId: 'mine', cellValue: { '1': { '1': { v: 'mine' } } } } }) }]
    const remote: OpEntry = { room: 'r', seq: 1, client_id: 'p', ops: rebaseUnit([srv({ '1': { '1': { v: 'theirs' } } })], 'mine') }
    expect(reconcileRemote(remote, owned)).toEqual([])
  })
})

describe('reconcileRemote', () => {
  const entry = (seq: number, cells: Record<string, Record<string, unknown>>): OpEntry => ({ room: 'r', seq, client_id: 'peer', ops: [srv(cells)] })

  it('lists the cells an op writes', () => {
    expect([...cellsOf(srv({ '1': { '2': { v: 1 }, '3': { v: 2 } } }))].length).toBe(2)
    expect(cellsOf({ id: 'sheet.mutation.insert-row', params: {} }).size).toBe(0)
    expect(cellsOf({ id: SET_RANGE_VALUES, params: { cellValue: null } }).size).toBe(0)
  })

  it('keeps my pending and later cells, applies the rest', () => {
    const owned = [
      { seq: null, cells: cellsOf(srv({ '1': { '1': { v: 'mine' } } })) },
      { seq: 9, cells: cellsOf(srv({ '2': { '2': { v: 'mine-later' } } })) },
      { seq: 3, cells: cellsOf(srv({ '3': { '3': { v: 'mine-earlier' } } })) },
    ]
    const remote = entry(5, { '1': { '1': { v: 'theirs' }, '2': { v: 'ok' } }, '2': { '2': { v: 'theirs' } }, '3': { '3': { v: 'wins' } } })
    const out = reconcileRemote(remote, owned)
    expect(out).toHaveLength(1)
    expect((out[0].params as { cellValue: unknown }).cellValue).toEqual({ '1': { '2': { v: 'ok' } }, '3': { '3': { v: 'wins' } } })
  })

  it('drops an op left with nothing, passes non-cell ops through, respects sheet identity', () => {
    const owned = [{ seq: null, cells: cellsOf(srv({ '1': { '1': { v: 1 } } })) }]
    expect(reconcileRemote(entry(1, { '1': { '1': { v: 2 } } }), owned)).toEqual([])
    const other = { room: 'r', seq: 1, client_id: 'p', ops: [srv({ '1': { '1': { v: 2 } } }, 's2'), { id: 'sheet.mutation.set-worksheet-name', params: { name: 'n' } }] }
    expect(reconcileRemote(other, owned)).toHaveLength(2)
    expect(reconcileRemote(entry(1, { '1': { '1': { v: 2 } } }), [])).toEqual(entry(1, { '1': { '1': { v: 2 } } }).ops)
  })

  it('prunes owned entries the log has passed', () => {
    const owned = [
      { seq: null, cells: new Set<string>() },
      { seq: 4, cells: new Set<string>() },
      { seq: 8, cells: new Set<string>() },
    ]
    expect(pruneOwned(owned, 6).map((o) => o.seq)).toEqual([null, 8])
  })

  it('tracks both absolute matrices of a move-range and filters later-owned cells', () => {
    const move: OpRecord = {
      id: MOVE_RANGE,
      params: {
        unitId: 'wb',
        fromRange: { startRow: 1, endRow: 1, startColumn: 0, endColumn: 1 },
        toRange: { startRow: 5, endRow: 5, startColumn: 0, endColumn: 1 },
        from: { subUnitId: 's1', value: { 1: { 0: null, 1: null } } },
        to: { subUnitId: 's1', value: { 5: { 0: { v: 'a' }, 1: { v: 'b' } } } },
      },
    }
    expect(cellsOf(move).size).toBe(4)
    const remote: OpEntry = { room: 'r', seq: 4, client_id: 'peer', ops: [move] }
    const owned = [{ seq: null, cells: cellsOf(srv({ 1: { 0: { v: 'local' } }, 5: { 1: { v: 'local' } } })) }]
    const [filtered] = reconcileRemote(remote, owned)
    expect((filtered.params.from as { value: unknown }).value).toEqual({ 1: { 1: null } })
    expect((filtered.params.to as { value: unknown }).value).toEqual({ 5: { 0: { v: 'a' } } })
  })

  it('splits a reordered range into atomic column bands around protected cells', () => {
    const reorder: OpRecord = {
      id: REORDER_RANGE,
      params: { unitId: 'wb', subUnitId: 's1', range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 2 }, order: { 0: 1, 1: 0 } },
    }
    expect(cellsOf(reorder).size).toBe(6)
    const remote: OpEntry = { room: 'r', seq: 4, client_id: 'peer', ops: [reorder] }
    const owned = [{ seq: null, cells: cellsOf(srv({ 0: { 1: { v: 'local' } } })) }]
    const filtered = reconcileRemote(remote, owned)
    expect(filtered.map((op) => op.params.range)).toEqual([
      { startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 },
      { startRow: 0, endRow: 1, startColumn: 2, endColumn: 2 },
    ])
    expect(filtered.map((op) => op.params.order)).toEqual([{ 0: 1, 1: 0 }, { 0: 1, 1: 0 }])
  })
})

describe('SubmitQueue', () => {
  it('batches a tick, ships in order, acks with seqs', async () => {
    const sent: number[][] = []
    let seq = 0
    const acked: number[] = []
    const q = new SubmitQueue(
      async (ops) => {
        sent.push(ops.map((o) => o.params.n as number))
        return ++seq
      },
      { onQueued: () => {}, onAcked: (_ops, s) => acked.push(s) },
    )
    q.push({ id: 'x', params: { n: 1 } })
    q.push({ id: 'x', params: { n: 2 } })
    await new Promise((r) => setTimeout(r, 0))
    q.push({ id: 'x', params: { n: 3 } })
    await new Promise((r) => setTimeout(r, 0))
    expect(sent).toEqual([[1, 2], [3]])
    expect(acked).toEqual([1, 2])
    expect(q.pending).toBe(0)
  })

  it('keeps a failed submission queued and retries on flush', async () => {
    let fail = true
    const sent: number[] = []
    const q = new SubmitQueue(
      async (ops) => {
        if (fail) throw new Error('down')
        sent.push(ops.length)
        return 1
      },
      { onQueued: () => {}, onAcked: () => {} },
    )
    q.push({ id: 'x', params: {} })
    await new Promise((r) => setTimeout(r, 0))
    expect(q.pending).toBe(1)
    fail = false
    await q.flush()
    expect(sent).toEqual([1])
    expect(q.pending).toBe(0)
  })
})

describe('SubmitQueue Stage C', () => {
  it('on STALE_BASE runs onStale and retries the (possibly transformed) head submission', async () => {
    const calls: string[] = []
    let stale = true
    const q = new SubmitQueue(
      async (ops) => {
        if (stale) {
          calls.push('stale:' + (ops[0].params.n as number))
          stale = false
          throw new Error('STALE_BASE: catch up and resubmit')
        }
        calls.push('ok:' + (ops[0].params.n as number))
        return 7
      },
      {
        onQueued: () => {},
        onAcked: (_o, seq) => calls.push('ack:' + seq),
        onStale: async () => {
          calls.push('catchup')
          q.transformQueued((ops) => ops.map((o) => ({ id: o.id, params: { n: (o.params.n as number) + 100 } })))
        },
      },
    )
    q.push({ id: 'x', params: { n: 1 } })
    await new Promise((r) => setTimeout(r, 10))
    expect(calls).toEqual(['stale:1', 'catchup', 'ok:101', 'ack:7'])
    expect(q.pending).toBe(0)
  })

  it('other errors keep the submission queued without onStale', async () => {
    let fails = 1
    const q = new SubmitQueue(
      async () => {
        if (fails-- > 0) throw new Error('down')
        return 1
      },
      { onQueued: () => {}, onAcked: () => {}, onStale: async () => { throw new Error('must not be called') } },
    )
    q.push({ id: 'x', params: {} })
    await new Promise((r) => setTimeout(r, 5))
    expect(q.pending).toBe(1)
    await q.flush()
    expect(q.pending).toBe(0)
  })

  it('transformQueued drops emptied submissions and allPending sees batch + queued', async () => {
    const q = new SubmitQueue(async () => Promise.reject(new Error('down')), { onQueued: () => {}, onAcked: () => {} })
    q.push({ id: 'a', params: { keep: false } })
    await new Promise((r) => setTimeout(r, 5))
    q.push({ id: 'b', params: { keep: true } })
    expect(q.allPending().map((o) => o.id)).toEqual(['a', 'b'])
    q.transformQueued((ops) => ops.filter((o) => o.params.keep === true))
    expect(q.allPending().map((o) => o.id)).toEqual(['b'])
    expect(q.pending).toBe(1) // batch only; the queued submission was emptied
    q.clear()
  })
})
