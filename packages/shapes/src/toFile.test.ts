import { describe, expect, it } from 'vitest'
import { toWireShapeRemove, toWireShapes, toWireShapeUpdate } from './toFile'
import type { ShapeSpec } from './types'

function spec(overrides: Partial<ShapeSpec> = {}): ShapeSpec {
  return { id: 's1', kind: 'rect', fill: '#eef4f2', stroke: '#0fa98f', strokeWidth: 1.5, ...overrides }
}

describe('toWireShapes', () => {
  it('carries the spec fields through with the active sheet name', () => {
    const { shapes, skipped } = toWireShapes([{ spec: spec({ text: 'Target' }) }], 'Data')
    expect(skipped).toEqual([])
    expect(shapes).toHaveLength(1)
    expect(shapes[0]).toMatchObject({ sheetName: 'Data', kind: 'rect', text: 'Target', fill: '#eef4f2', stroke: '#0fa98f' })
  })

  it('uses the file cell anchor when the shape has one', () => {
    const { shapes } = toWireShapes(
      [{ spec: spec(), cellAnchor: { FromCol: 3, FromRow: 4, ToCol: 7, ToRow: 8 } }],
      'Data',
    )
    expect(shapes[0].anchor).toEqual({ fromCol: 3, fromRow: 4, toCol: 7, toRow: 8 })
  })

  it('gives fresh (no-anchor) shapes distinct cascading default anchors', () => {
    const { shapes } = toWireShapes(
      [{ spec: spec({ id: 'a' }) }, { spec: spec({ id: 'b' }) }, { spec: spec({ id: 'c' }) }],
      'Data',
    )
    const anchors = shapes.map((s) => `${s.anchor.fromCol},${s.anchor.fromRow}`)
    expect(new Set(anchors).size).toBe(3) // no two land on the same cell
  })

  it('skips every shape (together) when the sheet is gone', () => {
    const { shapes, skipped } = toWireShapes([{ spec: spec({ text: 'Note' }) }], null)
    expect(shapes).toEqual([])
    expect(skipped).toEqual(['"Note": its sheet no longer exists'])
  })

  it('omits empty-string fields as undefined (matches SHAPE_DEFAULTS.text)', () => {
    const { shapes } = toWireShapes([{ spec: { id: 's', kind: 'text', text: 'hi', fill: '', stroke: '' } }], 'Data')
    expect(shapes[0].fill).toBeUndefined()
    expect(shapes[0].stroke).toBeUndefined()
  })

  it('refuses to serialize a hydrated shape as a duplicate add', () => {
    const native = spec({ nativeIdentity: { drawingPart: 'xl/drawings/drawing1.xml', objectId: 1 } })
    expect(toWireShapes([{ spec: native }], 'Data')).toEqual({
      shapes: [],
      skipped: ['shape s1: it already has native identity; use an update request'],
    })
  })
})

describe('native lifecycle wire requests', () => {
  const nativeIdentity = { drawingPart: 'xl/drawings/drawing3.xml', objectId: 42 }

  it('builds stable identity-bound remove and update requests', () => {
    const native = spec({ nativeIdentity, text: 'Native' })
    expect(toWireShapeRemove(native)).toEqual({
      request: { operation: 'remove', identity: nativeIdentity },
      skipped: [],
    })
    expect(toWireShapeUpdate({ spec: native, cellAnchor: { FromCol: 1, FromRow: 2, ToCol: 4, ToRow: 6 } }, 'Data')).toEqual({
      request: {
        operation: 'update',
        identity: nativeIdentity,
        shape: expect.objectContaining({ sheetName: 'Data', kind: 'rect', text: 'Native', anchor: { fromCol: 1, fromRow: 2, toCol: 4, toRow: 6 } }),
      },
      skipped: [],
    })
  })

  it('fails closed for browser-created or malformed identities', () => {
    expect(toWireShapeRemove(spec())).toEqual({ skipped: ['shape s1: it has no valid hydrated native identity'] })
    expect(toWireShapeUpdate({ spec: spec({ nativeIdentity: { drawingPart: '../drawing.xml', objectId: 1 } }) }, 'Data')).toEqual({
      skipped: ['shape s1: it has no valid hydrated native identity'],
    })
  })

  it('refuses a native update without its hydrated anchor', () => {
    expect(toWireShapeUpdate({ spec: spec({ nativeIdentity }) }, 'Data')).toEqual({
      skipped: ['shape s1: it has no hydrated native cell anchor'],
    })
  })

  it('does not turn a missing worksheet into an update', () => {
    expect(toWireShapeUpdate({ spec: spec({ nativeIdentity }) }, null)).toEqual({ skipped: ['"rect": its sheet no longer exists'] })
  })
})
