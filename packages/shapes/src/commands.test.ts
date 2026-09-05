import type { FUniver } from '@univerjs/core/lib/facade'
import { describe, expect, it } from 'vitest'
import { ShapeCommandController, type ShapeUndoRecord } from './commands'
import { ShapeManager } from './manager'
import type { ShapeSpec } from './types'

function harness(accept = true) {
  const disposed: string[] = []
  const sheet = {
    getSheetId: () => 'sheet-1',
    addFloatDomToPosition: (_config: unknown, id: string) => accept ? { id, dispose: () => disposed.push(id) } : null,
  }
  return { api: { getActiveWorkbook: () => ({ getId: () => 'book-1', getActiveSheet: () => sheet, getSheetBySheetId: (id: string) => id === 'sheet-1' ? sheet : null }) } as unknown as FUniver, disposed }
}

function spec(overrides: Partial<ShapeSpec> = {}): ShapeSpec {
  return { id: 'shape-1', kind: 'rect', text: 'Original', fill: '#ffffff', ...overrides }
}

describe('ShapeCommandController', () => {
  it('returns live handles and records complete add/update/remove inverses', () => {
    const host = harness()
    const records: ShapeUndoRecord[] = []
    const subject = new ShapeCommandController(new ShapeManager(host.api), { push: (record) => records.push(record) })
    const handle = subject.add(spec())
    expect(handle?.value?.text).toBe('Original')
    expect(handle?.update({ text: 'Updated' })).toBe(true)
    expect(handle?.value?.text).toBe('Updated')
    expect(handle?.remove()).toBe(true)
    expect(handle?.value).toBeUndefined()
    expect(records.map((record) => record.label)).toEqual(['Add shape', 'Update shape', 'Remove shape'])
    expect(records[0].before.shapes).toEqual([])
    expect(records[1].before.shapes[0].spec.text).toBe('Original')
    expect(records[2].after.shapes).toEqual([])
  })

  it('restores native identity and exact cell anchors', () => {
    const host = harness()
    const subject = new ShapeCommandController(new ShapeManager(host.api))
    const nativeIdentity = { drawingPart: 'xl/drawings/drawing1.xml', objectId: 7 }
    const anchor = { FromCol: 3, FromRow: 4, ToCol: 8, ToRow: 12 }
    subject.add(spec({ nativeIdentity }), anchor)
    const snapshot = subject.snapshot()
    subject.update('shape-1', { text: 'Changed' })
    expect(subject.restore(snapshot)).toBe(true)
    expect(subject.getById('shape-1')?.value).toMatchObject({ text: 'Original', nativeIdentity })
    expect(subject.snapshot().shapes[0].cellAnchor).toEqual(anchor)
  })

  it('creates catalog shapes and rejects invalid snapshots without mutation', () => {
    const host = harness()
    const subject = new ShapeCommandController(new ShapeManager(host.api))
    const created = subject.create('flowChartDecision')
    expect(created?.value).toMatchObject({ kind: 'flowChartDecision' })
    const before = subject.snapshot()
    expect(subject.restore({ version: 1, shapes: [{ spec: spec({ id: '', kind: 'rect' }), sheetId: 'sheet-1' }] })).toBe(false)
    expect(subject.snapshot()).toEqual(before)
    expect(subject.create('unknown' as never)).toBeNull()
  })

  it('protects kind and native identity from untrusted update params', () => {
    const host = harness()
    const subject = new ShapeCommandController(new ShapeManager(host.api))
    const nativeIdentity = { drawingPart: 'xl/drawings/drawing1.xml', objectId: 7 }
    subject.add(spec({ nativeIdentity }))
    expect(subject.update('shape-1', { id: 'other', kind: 'ellipse', nativeIdentity: { ...nativeIdentity, objectId: 9 }, text: 'Safe' } as never)).toBe(true)
    expect(subject.getById('shape-1')?.value).toMatchObject({ id: 'shape-1', kind: 'rect', text: 'Safe', nativeIdentity })
    expect(subject.update('shape-1', { strokeWidth: Number.NaN })).toBe(false)
    expect(subject.update('shape-1', { unknown: true } as never)).toBe(false)
    expect(subject.add(spec({ id: 'bad', nativeIdentity: { drawingPart: '../drawing.xml', objectId: 1 } }))).toBeNull()
  })

  it('attempts rollback when a snapshot cannot mount', () => {
    const rejecting = harness(false)
    const subject = new ShapeCommandController(new ShapeManager(rejecting.api))
    expect(subject.restore({ version: 1, shapes: [{ spec: spec(), sheetId: 'sheet-1' }] })).toBe(false)
    expect(subject.list()).toEqual([])
  })

  it('restores onto the recorded worksheet even when the active sheet changed', () => {
    const mounted: string[] = []
    const makeSheet = (id: string) => ({
      getSheetId: () => id,
      addFloatDomToPosition: () => { mounted.push(id); return { id: `dom-${id}`, dispose: () => {} } },
    })
    const sheets = { one: makeSheet('one'), two: makeSheet('two') }
    let active: keyof typeof sheets = 'one'
    const api = { getActiveWorkbook: () => ({
      getActiveSheet: () => sheets[active],
      getSheetBySheetId: (id: string) => sheets[id as keyof typeof sheets] ?? null,
    }) } as unknown as FUniver
    const subject = new ShapeCommandController(new ShapeManager(api))
    subject.add(spec())
    const snapshot = subject.snapshot()
    subject.remove('shape-1')
    active = 'two'
    expect(subject.restore(snapshot)).toBe(true)
    expect(subject.snapshot().shapes[0].sheetId).toBe('one')
    expect(mounted).toEqual(['one', 'one'])
  })
})
