import { toWireShapes } from '../../../packages/shapes/src/toFile'
import type { ShapeKind, ShapeSpec } from '../../../packages/shapes/src/types'

export function playgroundShapeWire(kind: ShapeKind, fill: string, stroke: string) {
  const spec: ShapeSpec = {
    id: 'playground-shape',
    kind,
    text: kind,
    fill,
    stroke,
    strokeWidth: 2,
  }
  return toWireShapes([{ spec }], 'Sheet1')
}
