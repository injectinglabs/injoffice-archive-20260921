import type { FUniver } from '@univerjs/core/lib/facade'
import { describe, expect, it } from 'vitest'
import {
  ChartCollaborationSession,
  type ChartCollaborationAuthorityRequest,
  type ChartCollaborationEvent,
  type ChartCollaborationIntent,
} from './collaboration'
import { ChartCommandController, type ChartUndoRecord } from './commands'
import { ChartManager, type ChartManagerOptions } from './manager'
import type { ChartSpec } from './types'

function harness(options: ChartManagerOptions = {}, undo?: ChartUndoRecord[]) {
  const sheet = {
    getRange: () => ({ getValues: () => [['Label', 'Value'], ['A', 1]] }),
    addFloatDomToPosition: (_config: unknown, id: string) => ({ id, dispose: () => undefined }),
    addFloatDomToRange: (_range: unknown, _config: unknown, _options: unknown, id: string) => ({ id, dispose: () => undefined }),
  }
  const workbook = {
    getSheetBySheetId: (id: string) => id === 'sheet-1' ? sheet : null,
  }
  const api = { getActiveWorkbook: () => workbook } as unknown as FUniver
  return new ChartCommandController(new ChartManager(api, undefined, options), undo ? { push: (record) => undo.push(record) } : undefined)
}

function spec(id: string, overrides: Partial<ChartSpec> = {}): ChartSpec {
  return {
    id,
    type: 'Column',
    title: id,
    range: { sheetId: 'sheet-1', startRow: 0, startColumn: 0, endRow: 2, endColumn: 1 },
    ...overrides,
  }
}

function event(intent: ChartCollaborationIntent, revision: number): ChartCollaborationEvent {
  return { ...intent, revision }
}

function authority(allowed = true) {
  const requests: ChartCollaborationAuthorityRequest[] = []
  const state = {
    revision: 0,
    allowed,
    requests,
    hook: {
      get revision() { return state.revision },
      canApply(request: ChartCollaborationAuthorityRequest) {
        requests.push(request)
        return state.allowed
      },
    },
  }
  return state
}

describe('ChartCollaborationSession', () => {
  it('applies per-object lifecycle events without adding remote work to local undo', () => {
    const undo: ChartUndoRecord[] = []
    const commands = harness({}, undo)
    commands.add(spec('chart-1'))
    undo.length = 0
    const access = authority()
    let sequence = 0
    const subject = new ChartCollaborationSession('book-1', commands, {
      authority: access.hook,
      operationId: () => `op-${++sequence}`,
    })
    const observed: string[] = []
    subject.onResult(() => { throw new Error('observer failure') })
    subject.onResult(({ code }) => observed.push(code))
    const identity = { part: 'xl/charts/chart2.xml', drawingPart: 'xl/drawings/drawing1.xml', objectId: 2 }

    const create = subject.createIntent('alice', spec('chart-2', { nativeIdentity: identity }))!
    expect(create.mutation).toMatchObject({ kind: 'create', baseObjectRevision: 0, baseLayerRevision: 0 })
    expect(subject.receive(event(create, 1))).toMatchObject({ ok: true, applied: true, consumed: true, revision: 1 })

    const update = subject.updateIntent('alice', 'chart-2', { title: 'Updated' })!
    expect(update.mutation).toMatchObject({ kind: 'update', baseObjectRevision: 1 })
    expect(subject.receive(event(update, 2)).code).toBe('applied')
    expect(commands.getById('chart-2')?.value).toMatchObject({ title: 'Updated', nativeIdentity: identity })

    const layer = subject.layerIntent('alice', 'chart-1', 'bringToFront')!
    expect(layer.mutation).toMatchObject({ baseObjectRevision: 0, baseLayerRevision: 1 })
    expect(subject.receive(event(layer, 3)).code).toBe('applied')
    expect(commands.list().map(({ id }) => id)).toEqual(['chart-2', 'chart-1'])

    const remove = subject.removeIntent('alice', 'chart-2')!
    expect(remove.mutation).toMatchObject({ baseObjectRevision: 2, baseLayerRevision: 3 })
    expect(subject.receive(event(remove, 4)).code).toBe('applied')
    expect(commands.getById('chart-2')).toBeNull()
    expect(subject.revisionState()).toEqual({
      version: 1,
      revision: 4,
      objects: [
        { chartId: 'chart-1', sheetId: 'sheet-1', revision: 3, deleted: false },
        { chartId: 'chart-2', sheetId: 'sheet-1', revision: 4, deleted: true },
      ],
      layers: [{ sheetId: 'sheet-1', revision: 4 }],
    })
    expect(undo).toEqual([])
    expect(access.requests.map(({ mutation }) => mutation.kind)).toEqual(['create', 'update', 'layer', 'remove'])
    expect(observed).toEqual(['applied', 'applied', 'applied', 'applied'])
  })

  it('deterministically consumes same-object conflicts while allowing a later revision', () => {
    const commands = harness()
    commands.add(spec('chart-1'))
    const access = authority()
    let sequence = 0
    const subject = new ChartCollaborationSession('book-1', commands, {
      authority: access.hook,
      operationId: () => `conflict-${++sequence}`,
    })
    const first = subject.updateIntent('alice', 'chart-1', { title: 'Alice' })!
    const concurrent = subject.updateIntent('bob', 'chart-1', { title: 'Bob' })!

    expect(subject.receive(event(first, 1)).code).toBe('applied')
    expect(subject.receive(event(concurrent, 2))).toMatchObject({ code: 'conflict', consumed: true, revision: 2 })
    expect(commands.getById('chart-1')?.value?.title).toBe('Alice')

    const later = subject.updateIntent('bob', 'chart-1', { title: 'Bob rebased' })!
    expect(later.mutation).toMatchObject({ baseObjectRevision: 1 })
    expect(subject.receive(event(later, 3)).code).toBe('applied')
    expect(commands.getById('chart-1')?.value?.title).toBe('Bob rebased')
  })

  it('uses one sheet-layer revision to serialize relational order changes', () => {
    const commands = harness()
    commands.add(spec('chart-1'))
    commands.add(spec('chart-2'))
    commands.add(spec('chart-3'))
    const access = authority()
    let sequence = 0
    const subject = new ChartCollaborationSession('book-1', commands, {
      authority: access.hook,
      operationId: () => `layer-${++sequence}`,
    })
    const first = subject.layerIntent('alice', 'chart-1', 'bringToFront')!
    const concurrent = subject.layerIntent('bob', 'chart-3', 'sendToBack')!

    expect(subject.receive(event(first, 1)).code).toBe('applied')
    expect(subject.receive(event(concurrent, 2))).toMatchObject({ code: 'conflict', consumed: true })
    expect(commands.list().map(({ id }) => id)).toEqual(['chart-2', 'chart-3', 'chart-1'])
    expect(subject.revisionState().layers).toEqual([{ sheetId: 'sheet-1', revision: 1 }])
  })

  it('enforces authority revisions, consumes denials, and requires resync after a gap', () => {
    const commands = harness()
    commands.add(spec('chart-1'))
    const access = authority(false)
    const resyncCodes: string[] = []
    let sequence = 0
    const subject = new ChartCollaborationSession('book-1', commands, {
      authority: access.hook,
      operationId: () => `auth-${++sequence}`,
      onResyncRequired: ({ code }) => resyncCodes.push(code),
    })
    const denied = subject.updateIntent('mallory', 'chart-1', { title: 'Denied' })!
    expect(subject.receive(event(denied, 1))).toMatchObject({ code: 'unauthorized', consumed: true, revision: 1 })
    expect(commands.getById('chart-1')?.value?.title).toBe('chart-1')

    access.allowed = true
    const staleAuthority = subject.updateIntent('alice', 'chart-1', { title: 'Allowed' })!
    access.revision = 1
    expect(subject.receive(event(staleAuthority, 2))).toMatchObject({ code: 'authority-revision', consumed: false, revision: 1 })
    expect(subject.desynchronized).toBe(true)
    expect(subject.updateIntent('alice', 'chart-1', { title: 'Blocked until resync' })).toBeNull()

    expect(subject.resync({
      version: 1,
      revision: 2,
      objects: [{ chartId: 'chart-1', sheetId: 'sheet-1', revision: 0, deleted: false }],
      layers: [{ sheetId: 'sheet-1', revision: 0 }],
    })).toBe(true)
    expect(subject.desynchronized).toBe(false)
    const afterResync = subject.updateIntent('alice', 'chart-1', { title: 'After resync' })!
    expect(subject.receive(event(afterResync, 4))).toMatchObject({ code: 'gap', consumed: false, revision: 2 })
    expect(subject.desynchronized).toBe(true)
    expect(resyncCodes).toEqual(['authority-revision', 'gap'])
  })

  it('rejects identity replacement and cross-sheet moves without corrupting revision state', () => {
    const commands = harness()
    commands.add(spec('chart-1', {
      nativeIdentity: { part: 'xl/charts/chart1.xml', drawingPart: 'xl/drawings/drawing1.xml', objectId: 1 },
    }))
    const access = authority()
    const subject = new ChartCollaborationSession('book-1', commands, { authority: access.hook, operationId: () => 'safe-op' })
    const valid = subject.updateIntent('alice', 'chart-1', { title: 'Safe' })!
    const identityMutation = {
      ...event(valid, 1),
      mutation: { ...valid.mutation, patch: { title: 'Unsafe', nativeIdentity: undefined } },
    }
    expect(subject.receive(identityMutation)).toMatchObject({ code: 'malformed', consumed: false, revision: 0 })

    const move = {
      ...event(valid, 1),
      mutation: { ...valid.mutation, patch: { range: { ...spec('chart-1').range, sheetId: 'sheet-2' } } },
    }
    expect(subject.receive(move)).toMatchObject({ code: 'conflict', consumed: true, revision: 1 })
    expect(commands.getById('chart-1')?.value).toMatchObject({ title: 'chart-1', range: { sheetId: 'sheet-1' } })
  })

  it('retains removal tombstones and fails closed when a visual layer host refuses', () => {
    const commands = harness({ layerHost: { setChartOrder: () => false } })
    commands.add(spec('chart-1'))
    commands.add(spec('chart-2'))
    const access = authority()
    let sequence = 0
    const subject = new ChartCollaborationSession('book-1', commands, {
      authority: access.hook,
      operationId: () => `host-${++sequence}`,
    })
    const layer = subject.layerIntent('alice', 'chart-1', 'bringToFront')!
    expect(subject.receive(event(layer, 1))).toMatchObject({ code: 'application-failed', consumed: false, revision: 0 })
    expect(subject.desynchronized).toBe(true)
    expect(commands.list().map(({ id }) => id)).toEqual(['chart-1', 'chart-2'])

    expect(subject.resync(subject.revisionState())).toBe(true)
    const remove = subject.removeIntent('alice', 'chart-2')!
    expect(subject.receive(event(remove, 1)).code).toBe('applied')
    const recreate = {
      protocol: 'injoffice.chart-collaboration/v1',
      workbookId: 'book-1',
      operationId: 'reuse',
      actorId: 'alice',
      authorityRevision: 0,
      revision: 2,
      mutation: { kind: 'create', chart: spec('chart-2'), baseObjectRevision: 0, baseLayerRevision: 1 },
    }
    expect(subject.receive(recreate)).toMatchObject({ code: 'conflict', consumed: true, revision: 2 })
    expect(commands.getById('chart-2')).toBeNull()
  })

  it('consumes a valid server-ordered boundary layer event as a deterministic no-op', () => {
    const commands = harness()
    commands.add(spec('chart-1'))
    commands.add(spec('chart-2'))
    const access = authority()
    const subject = new ChartCollaborationSession('book-1', commands, { authority: access.hook })
    const received = {
      protocol: 'injoffice.chart-collaboration/v1',
      workbookId: 'book-1',
      operationId: 'boundary',
      actorId: 'alice',
      authorityRevision: 0,
      revision: 1,
      mutation: {
        kind: 'layer', chartId: 'chart-1', sheetId: 'sheet-1', operation: 'sendToBack',
        baseObjectRevision: 0, baseLayerRevision: 0,
      },
    }
    expect(subject.receive(received)).toMatchObject({ code: 'no-op', ok: true, applied: false, consumed: true, revision: 1 })
    expect(subject.revisionState()).toMatchObject({
      objects: [{ chartId: 'chart-1', revision: 1 }, { chartId: 'chart-2', revision: 0 }],
      layers: [{ sheetId: 'sheet-1', revision: 1 }],
    })
  })
})
