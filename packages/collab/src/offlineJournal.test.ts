import { describe, expect, it, vi } from 'vitest'
import {
  DurableOutboundJournal,
  OUTBOUND_JOURNAL_PROTOCOL,
  OutboundJournalCancelledError,
  createJournalSubmitter,
  type JournalSubmitRequest,
  type OutboundJournalOptions,
  type OutboundJournalStorage,
  type OutboundJournalTransport,
  type RetryScheduler,
} from './offlineJournal'
import { SubmitQueue, type OpRecord } from './sync'

const op = (value: number): OpRecord => ({ id: 'sheet.mutation.test', params: { value } })
const drain = async (turns = 20): Promise<void> => { for (let index = 0; index < turns; index++) await Promise.resolve() }

class MemoryStorage implements OutboundJournalStorage {
  value: string | null = null
  writes = 0
  failWrite = new Set<number>()
  read(): string | null { return this.value }
  write(_key: string, value: string): void {
    this.writes++
    if (this.failWrite.has(this.writes)) throw new Error('disk full')
    this.value = value
  }
}

class ManualScheduler implements RetryScheduler {
  tasks: Array<{ delay: number; cancelled: boolean; task: () => void }> = []
  schedule(delay: number, task: () => void): { cancel(): void } {
    const item = { delay, cancelled: false, task }
    this.tasks.push(item)
    return { cancel: () => { item.cancelled = true } }
  }
  runNext(): void {
    const item = this.tasks.shift()
    if (item && !item.cancelled) item.task()
  }
}

function options(overrides: Partial<OutboundJournalOptions> = {}) {
  const storage = overrides.storage as MemoryStorage ?? new MemoryStorage()
  const scheduler = overrides.scheduler as ManualScheduler ?? new ManualScheduler()
  let seq = 0
  const sent: JournalSubmitRequest[] = []
  const transport: OutboundJournalTransport = overrides.transport ?? {
    resync: async () => ({ headSeq: seq, replay: 'safe' }),
    submit: async (request) => {
      sent.push(request)
      return { idempotencyKey: request.idempotencyKey, seq: ++seq }
    },
  }
  let key = 0
  return {
    config: {
      room: 'book', clientId: 'client-a', idFactory: () => `key-${++key}`,
      storage, scheduler, transport, ...overrides,
    } satisfies OutboundJournalOptions,
    scheduler,
    sent,
    storage,
  }
}

describe('DurableOutboundJournal', () => {
  it('persists offline entries and replays them in ordinal order after restart', async () => {
    const first = options()
    const journal = new DurableOutboundJournal(first.config)
    await journal.open()
    const abandoned = journal.submit([op(1)], 4)
    void abandoned.catch(() => undefined)
    await drain()
    expect(JSON.parse(first.storage.value!).entries).toMatchObject([{ idempotencyKey: '1:key-1', ordinal: 1, createdBaseSeq: 4 }])
    journal.dispose()
    await expect(abandoned).rejects.toBeInstanceOf(OutboundJournalCancelledError)

    const restarted = options({ storage: first.storage, idFactory: () => 'unused' })
    const restored = new DurableOutboundJournal(restarted.config)
    await restored.open()
    expect(restored.pending()).toHaveLength(1)
    await restored.connect()
    await drain()
    expect(restarted.sent.map((request) => request.idempotencyKey)).toEqual(['1:key-1'])
    expect(restored.state).toMatchObject({ pending: 0, lastServerSeq: 1, status: 'ready' })
  })

  it('submits new entries one at a time with stable keys and advancing bases', async () => {
    const subject = options({
      transport: {
        resync: async () => ({ headSeq: 5, replay: 'safe' }),
        submit: async (request) => {
          subject.sent.push(request)
          return { idempotencyKey: request.idempotencyKey, seq: request.baseSeq + 1 }
        },
      },
    })
    const journal = new DurableOutboundJournal(subject.config)
    await journal.open()
    await journal.connect()
    const first = journal.submit([op(1)], 0)
    const second = journal.submit([op(2)], 0)
    expect(await first).toBe(6)
    expect(await second).toBe(7)
    expect(subject.sent.map(({ idempotencyKey, baseSeq }) => [idempotencyKey, baseSeq])).toEqual([['1:key-1', 5], ['2:key-2', 6]])
    expect(JSON.parse(subject.storage.value!).entries).toEqual([])
  })

  it('holds a durably enqueued batch from transport until its receipt is released', async () => {
    const subject = options()
    const journal = new DurableOutboundJournal(subject.config)
    await journal.open()
    await journal.connect()

    const receipt = await journal.enqueue([op(1)], 0, { deferSend: true })
    await drain()
    expect(subject.sent).toEqual([])
    expect(journal.pending()).toMatchObject([{ idempotencyKey: receipt.idempotencyKey, ops: [op(1)] }])

    receipt.release()
    receipt.release()
    expect(await receipt.acknowledged).toBe(1)
    expect(subject.sent).toHaveLength(1)
  })

  it('uses deterministic exponential retry and deduplicates an ack after persistence failure', async () => {
    const storage = new MemoryStorage()
    storage.failWrite.add(5) // ack persistence after the first server success
    const scheduler = new ManualScheduler()
    const requests: JournalSubmitRequest[] = []
    const transport: OutboundJournalTransport = {
      resync: async () => ({ headSeq: requests.length ? 1 : 0, replay: 'safe' }),
      submit: async (request) => {
        requests.push(request)
        return { idempotencyKey: request.idempotencyKey, seq: 1, deduplicated: requests.length > 1 }
      },
    }
    const subject = options({ storage, scheduler, transport, retry: { initialDelayMs: 100, multiplier: 2, maxDelayMs: 500 } })
    const journal = new DurableOutboundJournal(subject.config)
    await journal.open()
    await journal.connect()
    const acknowledged = journal.submit([op(1)])
    await drain()
    expect(scheduler.tasks.map((task) => task.delay)).toEqual([100])
    expect(journal.pending()[0]).toMatchObject({ idempotencyKey: '1:key-1', attempts: 1 })
    scheduler.runNext()
    await drain()
    expect(await acknowledged).toBe(1)
    expect(requests.map((request) => request.idempotencyKey)).toEqual(['1:key-1', '1:key-1'])
  })

  it('backs off repeated transport failures without sleeping', async () => {
    let failures = 3
    const scheduler = new ManualScheduler()
    const subject = options({
      scheduler,
      retry: { initialDelayMs: 10, multiplier: 3, maxDelayMs: 50 },
      transport: {
        resync: async () => ({ headSeq: 0, replay: 'safe' }),
        submit: async (request) => {
          if (failures-- > 0) throw new Error('offline')
          return { idempotencyKey: request.idempotencyKey, seq: 1 }
        },
      },
    })
    const journal = new DurableOutboundJournal(subject.config)
    await journal.open()
    await journal.connect()
    const acknowledged = journal.submit([op(1)])
    await drain()
    expect(scheduler.tasks[0].delay).toBe(10)
    scheduler.runNext(); await drain()
    expect(scheduler.tasks[0].delay).toBe(30)
    scheduler.runNext(); await drain()
    expect(scheduler.tasks[0].delay).toBe(50)
    scheduler.runNext(); await drain()
    expect(await acknowledged).toBe(1)
    expect(journal.state.retryAttempt).toBe(0)
  })

  it('runs another resync before replaying an edit captured during resync', async () => {
    let finishFirst!: (result: { headSeq: number; replay: 'safe' }) => void
    let resyncs = 0
    const sent: JournalSubmitRequest[] = []
    const scheduler = new ManualScheduler()
    const subject = options({
      scheduler,
      transport: {
        resync: async () => {
          resyncs++
          if (resyncs === 1) return new Promise((resolve) => { finishFirst = resolve })
          return { headSeq: 3, replay: 'safe' }
        },
        submit: async (request) => { sent.push(request); return { idempotencyKey: request.idempotencyKey, seq: 4 } },
      },
    })
    const journal = new DurableOutboundJournal(subject.config)
    await journal.open()
    const connecting = journal.connect()
    await drain()
    const pending = journal.submit([op(1)])
    await drain()
    finishFirst({ headSeq: 3, replay: 'safe' })
    expect(await connecting).toBe(false)
    expect(journal.state.status).toBe('retry-wait')
    expect(sent).toEqual([])
    scheduler.runNext()
    await drain()
    expect(await pending).toBe(4)
    expect(resyncs).toBe(2)
  })

  it('blocks unsafe replay, permits cancellation, and can resync again', async () => {
    let safe = false
    const blocked = vi.fn()
    const subject = options({
      hooks: { onBlocked: blocked },
      transport: {
        resync: async () => ({ headSeq: 8, replay: safe ? 'safe' : 'blocked' }),
        submit: async (request) => ({ idempotencyKey: request.idempotencyKey, seq: 9 }),
      },
    })
    const journal = new DurableOutboundJournal(subject.config)
    await journal.open()
    const pending = journal.submit([op(1)])
    void pending.catch(() => undefined)
    await drain()
    expect(await journal.connect()).toBe(false)
    expect(journal.state.status).toBe('blocked')
    expect(blocked).toHaveBeenCalledWith('unsafe-replay', expect.anything())
    await expect(journal.submit([op(2)])).rejects.toThrow('blocked')
    safe = true
    expect(await journal.connect()).toBe(true)
    expect(await pending).toBe(9)
  })

  it('accepts only exact-key resync rewrites and submits transformed operations', async () => {
    const blocked = vi.fn()
    const sent: JournalSubmitRequest[] = []
    let valid = false
    const subject = options({
      hooks: { onBlocked: blocked },
      transport: {
        resync: async (request) => ({
          headSeq: 4,
          replay: 'safe',
          rebased: valid ? request.pending.map((entry) => ({ idempotencyKey: entry.idempotencyKey, ops: [op(99)] })) : [],
        }),
        submit: async (request) => { sent.push(request); return { idempotencyKey: request.idempotencyKey, seq: 5 } },
      },
    })
    const journal = new DurableOutboundJournal(subject.config)
    await journal.open()
    const pending = journal.submit([op(1)])
    await drain()
    expect(await journal.connect()).toBe(false)
    expect(blocked).toHaveBeenCalledWith('invalid-resync', expect.anything())
    valid = true
    expect(await journal.connect()).toBe(true)
    expect(await pending).toBe(5)
    expect(sent[0].ops).toEqual([op(99)])
  })

  it('blocks an old acknowledgement unless the server confirms deduplication', async () => {
    const blocked = vi.fn()
    const subject = options({
      hooks: { onBlocked: blocked },
      transport: {
        resync: async () => ({ headSeq: 5, replay: 'safe' }),
        submit: async (request) => ({ idempotencyKey: request.idempotencyKey, seq: 5 }),
      },
    })
    const journal = new DurableOutboundJournal(subject.config)
    await journal.open()
    await journal.connect()
    const pending = journal.submit([op(1)])
    void pending.catch(() => undefined)
    await drain()
    expect(journal.state).toMatchObject({ status: 'blocked', pending: 1, lastServerSeq: 5 })
    expect(blocked).toHaveBeenCalledWith('invalid-resync', expect.objectContaining({ seq: 5 }))
    const [entry] = journal.pending()
    await journal.cancel(entry.idempotencyKey)
    await expect(pending).rejects.toBeInstanceOf(OutboundJournalCancelledError)
  })

  it('disconnects by aborting in-flight work and retains it for reconnect', async () => {
    let submitCalls = 0
    const subject = options({
      transport: {
        resync: async () => ({ headSeq: 0, replay: 'safe' }),
        submit: (request, signal) => new Promise((resolve, reject) => {
          submitCalls++
          if (submitCalls > 1) return resolve({ idempotencyKey: request.idempotencyKey, seq: 1 })
          signal.addEventListener('abort', () => reject(new Error('aborted')), { once: true })
        }),
      },
    })
    const journal = new DurableOutboundJournal(subject.config)
    await journal.open()
    await journal.connect()
    const pending = journal.submit([op(1)])
    await drain()
    journal.disconnect()
    await drain()
    expect(journal.state).toMatchObject({ status: 'offline', pending: 1 })
    expect(subject.scheduler.tasks).toHaveLength(0)
    await journal.connect()
    expect(await pending).toBe(1)
  })

  it('rejects a submission whose storage commit finishes after disposal', async () => {
    const storage = new MemoryStorage()
    const subject = options({ storage })
    const journal = new DurableOutboundJournal(subject.config)
    await journal.open()
    let finish!: () => void
    storage.write = (_key, value) => new Promise<void>((resolve) => {
      finish = () => { storage.value = value; resolve() }
    })
    const pending = journal.submit([op(1)])
    void pending.catch(() => undefined)
    await drain()
    journal.dispose()
    finish()
    await expect(pending).rejects.toBeInstanceOf(OutboundJournalCancelledError)
    expect(journal.state.status).toBe('disposed')
  })

  it('ignores a late resync result after disposal', async () => {
    let finish!: (result: { headSeq: number; replay: 'safe' }) => void
    const subject = options({
      transport: {
        resync: () => new Promise((resolve) => { finish = resolve }),
        submit: async (request) => ({ idempotencyKey: request.idempotencyKey, seq: 1 }),
      },
    })
    const journal = new DurableOutboundJournal(subject.config)
    await journal.open()
    const connecting = journal.connect()
    await drain()
    journal.dispose()
    finish({ headSeq: 9, replay: 'safe' })
    expect(await connecting).toBe(false)
    expect(journal.state).toMatchObject({ status: 'disposed', lastServerSeq: 0 })
  })

  it('refuses corrupt, incompatible, and cross-client persisted journals', async () => {
    const cases = [
      ['{', 'corrupt'],
      [JSON.stringify({ protocol: 'injoffice.collab.outbound-journal/v2' }), 'version'],
      [JSON.stringify({ protocol: OUTBOUND_JOURNAL_PROTOCOL, room: 'book', clientId: 'other', lastServerSeq: 0, nextOrdinal: 1, entries: [] }), 'identity'],
    ] as const
    for (const [raw, code] of cases) {
      const storage = new MemoryStorage()
      storage.value = raw
      const subject = options({ storage })
      const journal = new DurableOutboundJournal(subject.config)
      await expect(journal.open()).rejects.toMatchObject({ code })
      expect(storage.value).toBe(raw)
    }
  })

  it('plugs into SubmitQueue through its public send boundary', async () => {
    const subject = options()
    const journal = new DurableOutboundJournal(subject.config)
    await journal.open()
    await journal.connect()
    const acked = vi.fn()
    const queue = new SubmitQueue(createJournalSubmitter(journal, () => journal.state.lastServerSeq), { onQueued: vi.fn(), onAcked: acked })
    queue.push(op(7))
    await drain(40)
    expect(acked).toHaveBeenCalledWith([op(7)], 1)
    expect(queue.pending).toBe(0)
  })

  it('keeps two clients ordered by a shared idempotent server', async () => {
    let serverSeq = 0
    const dedup = new Map<string, number>()
    const requests: string[] = []
    const transport: OutboundJournalTransport = {
      resync: async () => ({ headSeq: serverSeq, replay: 'safe' }),
      submit: async (request) => {
        requests.push(`${request.clientId}:${request.idempotencyKey}`)
        const previous = dedup.get(request.idempotencyKey)
        const seq = previous ?? ++serverSeq
        dedup.set(request.idempotencyKey, seq)
        return { idempotencyKey: request.idempotencyKey, seq, deduplicated: previous !== undefined }
      },
    }
    const aliceOptions = options({ clientId: 'alice-tab', idFactory: () => 'alice-1', transport })
    const bobOptions = options({ clientId: 'bob-tab', idFactory: () => 'bob-1', transport })
    const alice = new DurableOutboundJournal(aliceOptions.config)
    const bob = new DurableOutboundJournal(bobOptions.config)
    await Promise.all([alice.open(), bob.open()])
    await Promise.all([alice.connect(), bob.connect()])
    const [aliceSeq, bobSeq] = await Promise.all([alice.submit([op(1)]), bob.submit([op(2)])])
    expect([aliceSeq, bobSeq].sort()).toEqual([1, 2])
    expect(requests).toEqual(['alice-tab:1:alice-1', 'bob-tab:1:bob-1'])
    expect(alice.state.pending + bob.state.pending).toBe(0)
  })
})
