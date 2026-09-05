// Save-time OOXML serialization: turn mounted ShapeSpecs into the wire shape
// the gateway's xlsxpatch.AddShape consumes, so browser-created shapes
// survive Save as REAL workbook drawing objects (and round-trip back
// through the ?shapes=json bridge on the next load). Mirrors
// @injoffice/charts' toFile.ts.
//
// Unlike ChartSpec, a ShapeSpec carries no sheetId of its own — the same
// simplification ShapeManager already makes (every shape mounts on
// whichever sheet is active at insert time; there is no per-shape
// cross-sheet tracking yet). So every shape in one save is written against
// the SAME sheet: the one the host resolves as active.

import type { FileShapeAnchor } from './fromFile'
import type { NativeShapeIdentity, ShapeSpec } from './types'

/** Mirror of the gateway's wireShape (files_apply_specs.go). */
export interface WireShape {
  sheetName: string
  kind: string
  text?: string
  fill?: string
  stroke?: string
  strokeWidth?: number
  textColor?: string
  fontSize?: number
  anchor: { fromCol: number; fromRow: number; toCol: number; toRow: number }
}

export interface ToWireShapeInput {
  spec: ShapeSpec
  cellAnchor?: FileShapeAnchor
}

export interface ToWireShapeResult {
  shapes: WireShape[]
  skipped: string[]
}

export interface WireShapeRemove {
  operation: 'remove'
  identity: NativeShapeIdentity
}

export interface WireShapeUpdate {
  operation: 'update'
  identity: NativeShapeIdentity
  shape: WireShape
}

export interface ShapeLifecycleWireResult<T extends WireShapeRemove | WireShapeUpdate> {
  request?: T
  skipped: string[]
}

function nativeIdentityOf(spec: Pick<ShapeSpec, 'id' | 'nativeIdentity'>): NativeShapeIdentity | null {
  const identity = spec.nativeIdentity
  if (!identity || !/^xl\/drawings\/(?!_rels\/)[^/]+\.xml$/.test(identity.drawingPart)
    || !Number.isSafeInteger(identity.objectId) || identity.objectId <= 0 || identity.objectId > 0xffff_ffff) return null
  return { ...identity }
}

/** Build a fail-closed delete request. A browser-created shape cannot be
 * mistaken for a native object and silently converted into a no-op. */
export function toWireShapeRemove(spec: Pick<ShapeSpec, 'id' | 'nativeIdentity'>): ShapeLifecycleWireResult<WireShapeRemove> {
  const identity = nativeIdentityOf(spec)
  if (!identity) return { skipped: [`shape ${spec.id}: it has no valid hydrated native identity`] }
  return { request: { operation: 'remove', identity }, skipped: [] }
}

/** Build a one-object native update while retaining hydrated identity. */
export function toWireShapeUpdate(input: ToWireShapeInput, sheetName: string | null): ShapeLifecycleWireResult<WireShapeUpdate> {
  const identity = nativeIdentityOf(input.spec)
  if (!identity) return { skipped: [`shape ${input.spec.id}: it has no valid hydrated native identity`] }
  if (!sheetName) return { skipped: [`"${input.spec.text || input.spec.kind}": its sheet no longer exists`] }
  if (!input.cellAnchor) return { skipped: [`shape ${input.spec.id}: it has no hydrated native cell anchor`] }
  return { request: { operation: 'update', identity, shape: wireShape(input, sheetName, 0) }, skipped: [] }
}

/** Default placement for a shape that never had a file anchor (created in
 *  the browser): a 4×4-cell block, cascading diagonally so several new
 *  shapes in one save don't all land on top of each other. */
function defaultAnchor(index: number): WireShape['anchor'] {
  const step = index % 8
  const fromCol = 1 + step
  const fromRow = 1 + step
  return { fromCol, fromRow, toCol: fromCol + 4, toRow: fromRow + 4 }
}

function wireShape({ spec, cellAnchor }: ToWireShapeInput, sheetName: string, freshIndex: number): WireShape {
  const anchor = cellAnchor
    ? { fromCol: cellAnchor.FromCol, fromRow: cellAnchor.FromRow, toCol: cellAnchor.ToCol, toRow: cellAnchor.ToRow }
    : defaultAnchor(freshIndex)
  return {
    sheetName,
    kind: spec.kind,
    text: spec.text || undefined,
    fill: spec.fill || undefined,
    stroke: spec.stroke || undefined,
    strokeWidth: spec.strokeWidth,
    textColor: spec.textColor,
    fontSize: spec.fontSize,
    anchor,
  }
}

/** Convert fresh mounted shapes to gateway add payloads. Hydrated shapes are
 * refused so a host cannot duplicate them instead of issuing an update.
 * sheetName is the
 *  active sheet's display name (null when it's gone — every shape is
 *  skipped together in that case, since they share it). */
export function toWireShapes(inputs: ToWireShapeInput[], sheetName: string | null): ToWireShapeResult {
  if (!sheetName) {
    return { shapes: [], skipped: inputs.map(({ spec }) => `"${spec.text || spec.kind}": its sheet no longer exists`) }
  }
  const shapes: WireShape[] = []
  const skipped: string[] = []
  let freshIndex = 0
  for (const input of inputs) {
    if (input.spec.nativeIdentity) {
      skipped.push(`shape ${input.spec.id}: it already has native identity; use an update request`)
      continue
    }
    shapes.push(wireShape(input, sheetName, freshIndex++))
  }
  return { shapes, skipped }
}
