import { describe, expect, it } from 'vitest'
import { playgroundShapeWire } from './shapeWire'

describe('playground shape native wire', () => {
  it('emits a drawing-ml wire shape for a browser-created preset', () => {
    const result = playgroundShapeWire('roundRect', '#dff7ef', '#087f6b')
    expect(result.skipped).toEqual([])
    expect(result.shapes[0]).toMatchObject({
      sheetName: 'Sheet1',
      kind: 'roundRect',
      fill: '#dff7ef',
      stroke: '#087f6b',
    })
    expect(result.shapes[0]?.anchor).toMatchObject({ fromCol: 1, fromRow: 1 })
  })
})
