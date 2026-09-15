import { RenderCompileError, type RenderPathCommand } from './types.js'

/** Declared read-only policy string carried by connectors whose path comes from
 * the DrawingML preset catalog instead of the exact straight-line contract. */
export const CONNECTOR_PRESET_GEOMETRY_POLICY = 'pptx.connector-preset-preview'

type EvaluatedPath = { readonly path: readonly RenderPathCommand[]; readonly fillMode: string; readonly stroke: boolean }

/** Projects evaluated connector geometry onto the connector shaft. A connector
 * is exactly one open stroked path (moveTo followed by line/curve commands);
 * anything else is refused at compile time rather than painted as a nearby
 * shape. Commands are copied, never mutated. */
export function connectorGeometryPath(paths: readonly EvaluatedPath[], sourcePath: string): readonly RenderPathCommand[] {
  if (paths.length !== 1) throw new RenderCompileError('native.connectorGeometry', sourcePath, 'connector geometry requires exactly one path')
  const part = paths[0]!
  if (part.fillMode !== 'none' || !part.stroke) throw new RenderCompileError('native.connectorGeometry', sourcePath, 'connector paths must be stroked without fill')
  if (part.path.length < 2 || part.path[0]!.kind !== 'moveTo') throw new RenderCompileError('native.connectorGeometry', sourcePath, 'connector paths start with moveTo and one drawing command')
  for (let index = 1; index < part.path.length; index++) {
    const kind = part.path[index]!.kind
    if (kind !== 'lineTo' && kind !== 'cubicBezierTo' && kind !== 'quadBezierTo') {
      throw new RenderCompileError('native.connectorGeometry', `${sourcePath}.paths[0].commands[${index}]`, `connector paths are open polylines or curves, not ${kind}`)
    }
  }
  return part.path.map((command) => ({ ...command }))
}
