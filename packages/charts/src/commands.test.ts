import type { FUniver } from '@univerjs/core/lib/facade'
import { describe, expect, it } from 'vitest'
import { ChartCommandController, type ChartUndoRecord } from './commands'
import { ChartManager } from './manager'
import type { ChartSpec } from './types'

function harness() {
  const disposed: string[] = []
  const sheet = {
    getSheetId: () => 'sheet-1',
    getActiveRange: () => ({ getRange: () => ({ startRow: 1, startColumn: 2, endRow: 5, endColumn: 6 }) }),
    getRange: () => ({ getValues: () => [['Label', 'Value'], ['A', 1]] }),
    addFloatDomToPosition: (_config: unknown, id: string) => ({ id, dispose: () => disposed.push(id) }),
    addFloatDomToRange: (_range: unknown, _config: unknown, _options: unknown, id: string) => ({ id, dispose: () => disposed.push(id) }),
  }
  const workbook = {
    getId: () => 'book-1',
    getActiveSheet: () => sheet,
    getActiveRange: () => null,
    getSheetBySheetId: (id: string) => id === 'sheet-1' ? sheet : null,
  }
  return { api: { getActiveWorkbook: () => workbook } as unknown as FUniver, disposed }
}

function spec(overrides: Partial<ChartSpec> = {}): ChartSpec {
  return {
    id: 'chart-1',
    type: 'Column',
    title: 'Original',
    range: { sheetId: 'sheet-1', startRow: 0, startColumn: 0, endRow: 2, endColumn: 1 },
    ...overrides,
  }
}

describe('ChartCommandController', () => {
  it('returns live handles and records complete add/update/remove inverses', () => {
    const host = harness()
    const records: ChartUndoRecord[] = []
    const subject = new ChartCommandController(new ChartManager(host.api), { push: (record) => records.push(record) })
    const handle = subject.add(spec())
    expect(handle?.value?.title).toBe('Original')
    expect(handle?.update({ title: 'Updated' })).toBe(true)
    expect(handle?.value?.title).toBe('Updated')
    expect(handle?.remove()).toBe(true)
    expect(handle?.value).toBeUndefined()
    expect(records.map((record) => record.label)).toEqual(['Add chart', 'Update chart', 'Remove chart'])
    expect(records[0].before.charts).toEqual([])
    expect(records[1].before.charts[0].spec.title).toBe('Original')
    expect(records[2].after.charts).toEqual([])
  })

  it('restores native identity and exact cell anchors', () => {
    const host = harness()
    const subject = new ChartCommandController(new ChartManager(host.api))
    const nativeIdentity = { part: 'xl/charts/chart1.xml', drawingPart: 'xl/drawings/drawing1.xml', objectId: 7 }
    const anchor = { FromCol: 3, FromRow: 4, ToCol: 11, ToRow: 20 }
    subject.add(spec({ nativeIdentity }), anchor)
    const snapshot = subject.snapshot()
    subject.update('chart-1', { title: 'Changed' })
    expect(subject.restore(snapshot)).toBe(true)
    expect(subject.getById('chart-1')?.value).toMatchObject({ title: 'Original', nativeIdentity })
    expect(subject.snapshot().charts[0].cellAnchor).toEqual(anchor)
  })

  it('creates from the active selection and refuses invalid snapshots without mutation', () => {
    const host = harness()
    const subject = new ChartCommandController(new ChartManager(host.api))
    const created = subject.createFromSelection('Line', 'Selection')
    expect(created?.value).toMatchObject({
      type: 'Line',
      title: 'Selection',
      range: { sheetId: 'sheet-1', startRow: 1, startColumn: 2, endRow: 5, endColumn: 6 },
    })
    const before = subject.snapshot()
    expect(subject.restore({ version: 1, charts: [{ spec: spec({ id: '' }) }] })).toBe(false)
    expect(subject.snapshot()).toEqual(before)
  })

  it('rolls back when a snapshot references a missing sheet', () => {
    const host = harness()
    const subject = new ChartCommandController(new ChartManager(host.api))
    subject.add(spec())
    expect(subject.restore({ version: 1, charts: [{ spec: spec({ range: { ...spec().range, sheetId: 'missing' } }) }] })).toBe(false)
    expect(subject.getById('chart-1')?.value?.title).toBe('Original')
    expect(host.disposed).toContain('chart-1')
  })

  it('rejects invalid updates and cannot replace stable identity through untrusted params', () => {
    const host = harness()
    const subject = new ChartCommandController(new ChartManager(host.api))
    const identity = { part: 'xl/charts/chart1.xml', drawingPart: 'xl/drawings/drawing1.xml', objectId: 7 }
    subject.add(spec({ nativeIdentity: identity }))
    expect(subject.update('chart-1', { type: 'bogus' } as never)).toBe(false)
    expect(subject.update('chart-1', { id: 'other', nativeIdentity: { ...identity, objectId: 9 }, title: 'Safe' } as never)).toBe(true)
    expect(subject.getById('chart-1')?.value).toMatchObject({ id: 'chart-1', title: 'Safe', nativeIdentity: identity })
    expect(subject.add(spec({ id: 'bad', nativeIdentity: { ...identity, drawingPart: '../drawing.xml' } }))).toBeNull()
  })

  it('layers live handles with snapshot undo and refuses invalid operations', () => {
    const host = harness()
    const records: ChartUndoRecord[] = []
    const subject = new ChartCommandController(new ChartManager(host.api), { push: (record) => records.push(record) })
    const first = subject.add(spec())!
    subject.add(spec({ id: 'chart-2' }))
    subject.add(spec({ id: 'chart-3' }))
    records.length = 0

    expect(first.layer('bringToFront')).toBe(true)
    expect(subject.list().map(({ id }) => id)).toEqual(['chart-2', 'chart-3', 'chart-1'])
    expect(records).toHaveLength(1)
    expect(records[0]).toMatchObject({
      label: 'Layer chart: bringToFront',
      before: { charts: [{ spec: { id: 'chart-1' } }, { spec: { id: 'chart-2' } }, { spec: { id: 'chart-3' } }] },
      after: { charts: [{ spec: { id: 'chart-2' } }, { spec: { id: 'chart-3' } }, { spec: { id: 'chart-1' } }] },
    })
    expect(subject.layer('chart-1', 'bogus' as never)).toBe(false)
    expect(subject.layer('missing', 'sendToBack')).toBe(false)
    expect(records).toHaveLength(1)
  })
})
