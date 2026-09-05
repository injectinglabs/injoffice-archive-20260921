import { describe, expect, it } from 'vitest'
import { specFromFileShape, specsFromFileShapes, type FileShapeInfo } from './fromFile'

function info(overrides: Partial<FileShapeInfo> = {}): FileShapeInfo {
  return {
    identity: { drawingPart: 'xl/drawings/drawing1.xml', objectId: 7 },
    sheetName: 'Data',
    kind: 'rect',
    text: 'Target',
    fill: '#eef4f2',
    stroke: '#0fa98f',
    strokeWidth: 1.5,
    anchor: { FromCol: 3, FromRow: 4, ToCol: 7, ToRow: 8 },
    ...overrides,
  }
}

describe('specFromFileShape', () => {
  it('converts a known kind on a known sheet', () => {
    const conv = specFromFileShape(info(), { Data: 'sheet-1' })
    expect(conv).not.toBeNull()
    expect(conv!.sheetId).toBe('sheet-1')
    expect(conv!.cellAnchor).toEqual({ FromCol: 3, FromRow: 4, ToCol: 7, ToRow: 8 })
    expect(conv!.spec).toMatchObject({ kind: 'rect', text: 'Target', fill: '#eef4f2', stroke: '#0fa98f' })
    expect(conv!.spec.id).toBe('native-shape-xl_drawings_drawing1_xml-7')
    expect(conv!.spec.nativeIdentity).toEqual({ drawingPart: 'xl/drawings/drawing1.xml', objectId: 7 })
  })

  it('rejects an unrenderable kind', () => {
    expect(specFromFileShape(info({ kind: 'notARealPreset' }), { Data: 'sheet-1' })).toBeNull()
  })

  it('rejects a shape whose sheet is not in this workbook', () => {
    expect(specFromFileShape(info({ sheetName: 'Gone' }), { Data: 'sheet-1' })).toBeNull()
  })

  it('identity-derived ids are deterministic and distinct', () => {
    const a = specFromFileShape(info(), { Data: 's1' })!
    const b = specFromFileShape(info({ identity: { drawingPart: 'xl/drawings/drawing1.xml', objectId: 8 } }), { Data: 's1' })!
    expect(a.spec.id).not.toBe(b.spec.id)
    expect(specFromFileShape(info(), { Data: 's1' })!.spec.id).toBe(a.spec.id)
  })

  it('rejects missing and malformed native identities', () => {
    expect(specFromFileShape(info({ identity: undefined as never }), { Data: 's1' })).toBeNull()
    expect(specFromFileShape(info({ identity: { drawingPart: '../drawing.xml', objectId: 7 } }), { Data: 's1' })).toBeNull()
    expect(specFromFileShape(info({ identity: { drawingPart: 'xl/drawings/drawing1.xml', objectId: 0 } }), { Data: 's1' })).toBeNull()
  })
})

describe('specsFromFileShapes', () => {
  it('counts unconvertible shapes as skipped rather than dropping silently', () => {
    const { conversions, skipped } = specsFromFileShapes(
      [info(), info({ kind: 'notARealPreset' }), info({ sheetName: 'Gone' })],
      { Data: 'sheet-1' },
    )
    expect(conversions).toHaveLength(1)
    expect(skipped).toBe(2)
  })

  it('rejects duplicate native identities as ambiguous', () => {
    const { conversions, skipped } = specsFromFileShapes([info(), info({ text: 'Duplicate' })], { Data: 'sheet-1' })
    expect(conversions).toHaveLength(1)
    expect(skipped).toBe(1)
  })
})
