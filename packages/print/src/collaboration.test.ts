import { describe, expect, it, vi } from 'vitest'
import { PrintConfigurationCommandController, type PrintConfigurationUndoRecord } from './commands'
import {
  PRINT_COLLABORATION_PROTOCOL,
  PrintCollaborationSession,
  applyPrintCollaborationOperation,
  fingerprintPrintConfiguration,
  type PrintCollaborationOperation,
  type PrintCollaborationTransport,
} from './collaboration'
import { PrintManager } from './manager'
import type { PrintConfigurationSnapshotV1 } from './types'

function controller(save: (snapshot: Readonly<PrintConfigurationSnapshotV1>) => Promise<void> | void = () => {}) {
  const undo: PrintConfigurationUndoRecord[] = []
  const commands = new PrintConfigurationCommandController(
    new PrintManager({ print: vi.fn() }, 'sheet-1'),
    { save },
    { push: (record) => undo.push(record) },
  )
  return { commands, undo }
}

const operation = (before: PrintConfigurationSnapshotV1, after: PrintConfigurationSnapshotV1, overrides: Partial<PrintCollaborationOperation> = {}): PrintCollaborationOperation => ({
  protocol: PRINT_COLLABORATION_PROTOCOL,
  opId: 'op-1',
  clientId: 'client-a',
  kind: 'replace',
  action: 'replace',
  expectedFingerprint: fingerprintPrintConfiguration(before),
  value: after,
  ...overrides,
})

class Hub {
  snapshot: PrintConfigurationSnapshotV1
  sequence = 0
  private handlers = new Map<string, (entry: unknown) => void>()

  constructor(snapshot: PrintConfigurationSnapshotV1) { this.snapshot = structuredClone(snapshot) }

  transport(clientId: string): PrintCollaborationTransport {
    return {
      submit: async (room, value, baseSequence) => {
        if (baseSequence !== this.sequence) throw Object.assign(new Error('base revision is stale'), { code: 'STALE_BASE' })
        this.snapshot = applyPrintCollaborationOperation(this.snapshot, value)
        const entry = { room, operation: value, sequence: ++this.sequence }
        queueMicrotask(() => { for (const [id, handler] of this.handlers) if (id !== clientId) handler(structuredClone(entry)) })
        return structuredClone(entry)
      },
      subscribe: (_room, handler) => { this.handlers.set(clientId, handler); return () => { this.handlers.delete(clientId) } },
    }
  }
}

function session(commands: PrintConfigurationCommandController, transport: PrintCollaborationTransport, clientId: string, initialSequence = 0, options: Record<string, unknown> = {}) {
  let id = 0
  return new PrintCollaborationSession(commands, transport, {
    room: 'book-1', clientId, initialSequence, idFactory: () => `${clientId}-${++id}`, ...options,
  })
}

describe('print collaboration reducer', () => {
  it('replaces a fingerprinted complete configuration without mutating input', () => {
    const { commands } = controller()
    const before = commands.snapshot()
    const after = structuredClone(before)
    after.layout.direction = 'Landscape'
    after.render.gridlines = true
    const result = applyPrintCollaborationOperation(before, operation(before, after))

    expect(result).toEqual(after)
    expect(result).not.toBe(after)
    expect(before.layout.direction).toBe('Portrait')
    expect(() => applyPrintCollaborationOperation(after, operation(before, { ...after, render: { ...after.render, headings: true } })))
      .toThrowError(expect.objectContaining({ code: 'conflict' }))
  })

  it('canonicalizes fingerprints and rejects widened nested payloads', () => {
    const { commands } = controller()
    const before = commands.snapshot()
    const reordered = { render: before.render, version: 1 as const, layout: before.layout }
    expect(fingerprintPrintConfiguration(before)).toBe(fingerprintPrintConfiguration(reordered))
    expect(() => applyPrintCollaborationOperation(before, {
      ...operation(before, before), privileged: true,
    } as PrintCollaborationOperation)).toThrowError(expect.objectContaining({ code: 'invalid' }))
    expect(() => applyPrintCollaborationOperation(before, operation(before, {
      ...before, layout: { ...before.layout, pageSizeCustom: { width: 10, height: 10, unsafe: 1 } },
    } as PrintConfigurationSnapshotV1))).toThrowError(expect.objectContaining({ code: 'invalid' }))
  })
})

describe('PrintCollaborationSession', () => {
  it('converges two clients from server-ordered persistence-first commands', async () => {
    const first = controller()
    const second = controller()
    const hub = new Hub(first.commands.snapshot())
    const left = session(first.commands, hub.transport('left'), 'left')
    const right = session(second.commands, hub.transport('right'), 'right')
    left.start(); right.start()

    await expect(left.updateLayout({ direction: 'Landscape', margin: 'Narrow' })).resolves.toBe(true)
    await vi.waitFor(() => expect(second.commands.snapshot().layout.direction).toBe('Landscape'))
    await expect(right.updateRender({ gridlines: true, headings: true })).resolves.toBe(true)
    await vi.waitFor(() => expect(first.commands.snapshot().render.gridlines).toBe(true))

    expect(first.commands.snapshot()).toEqual(hub.snapshot)
    expect(second.commands.snapshot()).toEqual(hub.snapshot)
    expect(first.undo).toHaveLength(1)
    expect(first.undo[0]).toMatchObject({ label: 'Update print layout', before: { layout: { direction: 'Portrait' } }, after: { layout: { direction: 'Landscape' } } })
    expect(second.undo).toHaveLength(1)
    expect(second.undo[0].label).toBe('Update print rendering')
    expect(left.state.sequence).toBe(2)
    expect(right.state.sequence).toBe(2)
  })

  it('does not expose local changes before the authoritative acknowledgement', async () => {
    const { commands, undo } = controller()
    let release!: () => void
    const gate = new Promise<void>((resolve) => { release = resolve })
    const base = commands.snapshot()
    const transport: PrintCollaborationTransport = {
      submit: async (room, value) => {
        await gate
        return { room, operation: value, sequence: 1 }
      },
      subscribe: () => () => undefined,
    }
    const subject = session(commands, transport, 'local')
    const pending = subject.updateLayout({ direction: 'Landscape' })
    await Promise.resolve()
    expect(commands.snapshot()).toEqual(base)
    expect(undo).toEqual([])
    release()
    await expect(pending).resolves.toBe(true)
    expect(commands.snapshot().layout.direction).toBe('Landscape')
    expect(undo).toHaveLength(1)
  })

  it('serializes concurrent local patches from the latest acknowledged state', async () => {
    const { commands, undo } = controller()
    const hub = new Hub(commands.snapshot())
    const subject = session(commands, hub.transport('local'), 'local')
    await expect(Promise.all([
      subject.updateLayout({ direction: 'Landscape' }),
      subject.updateLayout({ paperSize: 'Letter' }),
    ])).resolves.toEqual([true, true])
    expect(commands.snapshot().layout).toMatchObject({ direction: 'Landscape', paperSize: 'Letter' })
    expect(hub.snapshot).toEqual(commands.snapshot())
    expect(subject.state.sequence).toBe(2)
    expect(undo).toHaveLength(2)
  })

  it('enforces local and received permission policy without partial application', async () => {
    const { commands, undo } = controller()
    const submit = vi.fn()
    const denied = session(commands, { submit, subscribe: () => () => undefined }, 'denied', 0, {
      authorize: () => false,
    })
    await expect(denied.updateLayout({ direction: 'Landscape' })).rejects.toMatchObject({ code: 'permission' })
    expect(submit).not.toHaveBeenCalled()
    expect(commands.snapshot().layout.direction).toBe('Portrait')
    expect(undo).toEqual([])

    const remote = session(commands, { submit, subscribe: () => () => undefined }, 'receiver', 0, {
      authorize: (_operation: PrintCollaborationOperation, context: { direction: string }) => context.direction === 'submit',
    })
    const next = commands.snapshot(); next.render.gridlines = true
    await expect(remote.receive({ room: 'book-1', sequence: 1, operation: operation(commands.snapshot(), next, { clientId: 'other' }) })).resolves.toBe('blocked')
    expect(remote.state).toMatchObject({ sequence: 0, blocked: true })
    expect(commands.snapshot().render.gridlines).toBe(false)
  })

  it('blocks stale bases and gaps, then persists an explicit resync without undo', async () => {
    const saved: PrintConfigurationSnapshotV1[] = []
    const { commands, undo } = controller((snapshot) => { saved.push(structuredClone(snapshot)) })
    const hub = new Hub(commands.snapshot())
    const seed = operation(hub.snapshot, { ...hub.snapshot, layout: { ...hub.snapshot.layout, paperSize: 'Letter' } })
    hub.snapshot = applyPrintCollaborationOperation(hub.snapshot, seed)
    hub.sequence = 1
    const resync = vi.fn()
    const subject = session(commands, hub.transport('late'), 'late', 0, { onResyncRequired: resync })

    await expect(subject.updateLayout({ direction: 'Landscape' })).rejects.toMatchObject({ code: 'stale-base' })
    expect(subject.state.blocked).toBe(true)
    expect(resync).toHaveBeenCalledWith(1, -1, expect.anything())
    await expect(subject.resync(hub.snapshot, hub.sequence)).resolves.toBe(true)
    expect(commands.snapshot().layout.paperSize).toBe('Letter')
    expect(saved).toHaveLength(1)
    expect(undo).toEqual([])

    const later = commands.snapshot(); later.render.headings = true
    await expect(subject.receive({ room: 'book-1', sequence: 3, operation: operation(commands.snapshot(), later, { clientId: 'other' }) })).resolves.toBe('blocked')
    expect(subject.state).toMatchObject({ sequence: 1, blocked: true })
  })

  it('queues resynchronization behind an in-flight acknowledgement', async () => {
    const saved: PrintConfigurationSnapshotV1[] = []
    const { commands, undo } = controller((snapshot) => { saved.push(structuredClone(snapshot)) })
    let release!: () => void
    const gate = new Promise<void>((resolve) => { release = resolve })
    const transport: PrintCollaborationTransport = {
      submit: async (room, value) => {
        await gate
        return { room, sequence: 1, operation: value }
      },
      subscribe: () => () => undefined,
    }
    const subject = session(commands, transport, 'local')
    const local = subject.updateLayout({ direction: 'Landscape' })
    const authoritative = commands.snapshot()
    authoritative.layout.paperSize = 'Letter'
    const resync = subject.resync(authoritative, 2)

    await Promise.resolve()
    expect(commands.snapshot().layout).toMatchObject({ direction: 'Portrait', paperSize: 'A4' })
    release()
    await expect(local).resolves.toBe(true)
    await expect(resync).resolves.toBe(true)
    expect(commands.snapshot().layout).toMatchObject({ direction: 'Portrait', paperSize: 'Letter' })
    expect(subject.state.sequence).toBe(2)
    expect(saved).toHaveLength(2)
    expect(undo).toHaveLength(1)
  })

  it('blocks after a resync adapter reports success without applying the complete snapshot', async () => {
    const { commands } = controller()
    const submit = vi.fn()
    const subject = session(commands, { submit, subscribe: () => () => undefined }, 'local', 0, {
      applySnapshot: (snapshot: PrintConfigurationSnapshotV1) => {
        commands.manager.updatePrintConfig({ direction: snapshot.layout.direction })
        return true
      },
    })
    const authoritative = commands.snapshot()
    authoritative.layout.direction = 'Landscape'
    authoritative.layout.paperSize = 'Letter'

    await expect(subject.resync(authoritative, 4)).resolves.toBe(false)
    expect(commands.snapshot().layout).toMatchObject({ direction: 'Landscape', paperSize: 'A4' })
    expect(subject.state).toEqual({ sequence: 0, blocked: true, disposed: false })
    await expect(subject.updateRender({ gridlines: true })).rejects.toMatchObject({ code: 'blocked' })
    expect(submit).not.toHaveBeenCalled()
  })

  it('routes undo restore through the server without recursively adding undo', async () => {
    const { commands, undo } = controller()
    const hub = new Hub(commands.snapshot())
    const subject = session(commands, hub.transport('local'), 'local')
    await subject.updateLayout({ direction: 'Landscape' })
    const original = undo[0].before
    undo.length = 0

    await expect(subject.restore(original)).resolves.toBe(true)
    expect(commands.snapshot().layout.direction).toBe('Portrait')
    expect(hub.snapshot.layout.direction).toBe('Portrait')
    expect(undo).toEqual([])
  })

  it('blocks altered acknowledgements and atomic persistence failures', async () => {
    const first = controller()
    const changedAck = session(first.commands, {
      submit: async (room, value) => ({ room, sequence: 1, operation: { ...value, action: 'replace' } }),
      subscribe: () => () => undefined,
    }, 'local')
    await expect(changedAck.updateLayout({ direction: 'Landscape' })).rejects.toMatchObject({ code: 'conflict' })
    expect(changedAck.state.blocked).toBe(true)
    expect(first.commands.snapshot().layout.direction).toBe('Portrait')

    const failure = new Error('durable store unavailable')
    const second = controller(() => { throw failure })
    const applyFailure = session(second.commands, {
      submit: async (room, value) => ({ room, sequence: 1, operation: value }),
      subscribe: () => () => undefined,
    }, 'local')
    await expect(applyFailure.updateRender({ gridlines: true })).rejects.toMatchObject({ code: 'conflict' })
    expect(applyFailure.state.blocked).toBe(true)
    expect(second.commands.snapshot().render.gridlines).toBe(false)
    expect(second.undo).toEqual([])
  })

  it('refuses malformed commands, isolates observers, and disposes subscriptions', async () => {
    const { commands } = controller()
    const dispose = vi.fn()
    const events: string[] = []
    const subject = session(commands, {
      submit: async (room, value) => ({ room, sequence: 1, operation: value }),
      subscribe: () => dispose,
    }, 'local', 0, { onEvent: (event: { type: string }) => { events.push(event.type); throw new Error('observer') } })
    subject.start(); subject.start()
    await expect(subject.updateLayout(null as never)).resolves.toBe(false)
    await expect(subject.replace({ version: 2 })).resolves.toBe(false)
    await expect(subject.updateRender({ gridlines: true })).resolves.toBe(true)
    expect(events).toEqual(['applied', 'submitted'])
    subject.dispose(); subject.dispose()
    expect(dispose).toHaveBeenCalledOnce()
    await expect(subject.updateLayout({ direction: 'Landscape' })).rejects.toMatchObject({ code: 'blocked' })
  })
})
