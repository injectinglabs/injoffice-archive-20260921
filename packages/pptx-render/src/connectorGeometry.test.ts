import { describe, expect, it } from 'vitest'
import { CONNECTOR_PRESET_GEOMETRY_POLICY, connectorGeometryPath } from './connectorGeometry.js'
import { RenderCompileError, type RenderPathCommand } from './types.js'

const elbow: RenderPathCommand[] = [{ kind: 'moveTo', x: 0, y: 0 }, { kind: 'lineTo', x: 1000, y: 0 }, { kind: 'lineTo', x: 1000, y: 500 }]
const curve: RenderPathCommand[] = [{ kind: 'moveTo', x: 0, y: 0 }, { kind: 'cubicBezierTo', x1: 500, y1: 0, x2: 500, y2: 500, x: 1000, y: 500 }]

describe('connectorGeometryPath', () => {
  it('copies a single open stroked polyline or curve without mutating the evaluated commands', () => {
    for (const commands of [elbow, curve]) {
      const paths = [{ path: commands, fillMode: 'none', stroke: true }]
      const before = JSON.stringify(paths)
      const result = connectorGeometryPath(paths, '$.elements.c.geometry')
      expect(result).toEqual(commands)
      expect(result[0]).not.toBe(commands[0])
      expect(JSON.stringify(paths)).toBe(before)
    }
    expect(CONNECTOR_PRESET_GEOMETRY_POLICY).toBe('pptx.connector-preset-preview')
  })

  it('refuses geometry that could paint a filled, closed, multi-path, or arc shape as a connector', () => {
    const cases: readonly { readonly path: readonly RenderPathCommand[]; readonly fillMode: string; readonly stroke: boolean }[][] = [
      [],
      [{ path: elbow, fillMode: 'none', stroke: true }, { path: elbow, fillMode: 'none', stroke: true }],
      [{ path: elbow, fillMode: 'norm', stroke: true }],
      [{ path: elbow, fillMode: 'none', stroke: false }],
      [{ path: [{ kind: 'moveTo', x: 0, y: 0 }], fillMode: 'none', stroke: true }],
      [{ path: [{ kind: 'lineTo', x: 0, y: 0 }, { kind: 'lineTo', x: 1, y: 1 }], fillMode: 'none', stroke: true }],
      [{ path: [...elbow, { kind: 'close' }], fillMode: 'none', stroke: true }],
      [{ path: [...elbow, { kind: 'arcTo', x: 0, y: 0, rx: 5, ry: 5, largeArc: false, clockwise: true }], fillMode: 'none', stroke: true }],
      [{ path: [...elbow, { kind: 'moveTo', x: 5, y: 5 }], fillMode: 'none', stroke: true }],
    ]
    for (const paths of cases) {
      expect(() => connectorGeometryPath(paths, '$.elements.c.geometry')).toThrow(RenderCompileError)
      try { connectorGeometryPath(paths, '$.elements.c.geometry') } catch (error) { expect((error as RenderCompileError).code).toBe('native.connectorGeometry') }
    }
  })
})
