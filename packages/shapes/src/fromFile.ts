import type { NativeShapeIdentity, ShapeKind, ShapeSpec } from './types'
import { SHAPE_KINDS } from './types'

// Bridge from the FILE side: the gateway's xlsxpatch reader (Go) summarizes
// a workbook's inline drawing shapes as ShapeInfo JSON (kind/text/styling +
// its own cell anchor — no separate anchor lookup needed, unlike a chart,
// since a shape carries no reference to another part); this module turns
// that into mountable ShapeSpecs. Mirrors @injoffice/charts' fromFile.ts.

/** Mirror of Go xlsxpatch.ShapeInfo (JSON tags). */
export interface FileShapeInfo {
  identity: NativeShapeIdentity
  sheetName: string
  kind: string
  text?: string
  fill?: string
  stroke?: string
  strokeWidth?: number
  textColor?: string
  fontSize?: number
  anchor: FileShapeAnchor
}

/** Mirror of Go xlsxpatch.ShapeAnchor. */
export interface FileShapeAnchor {
  FromCol: number
  FromRow: number
  ToCol: number
  ToRow: number
}

const RENDERABLE: ReadonlySet<string> = new Set(SHAPE_KINDS)

export interface FileShapeConversion {
  spec: ShapeSpec
  cellAnchor: FileShapeAnchor
  sheetId: string
}

function validIdentity(identity: NativeShapeIdentity | undefined): identity is NativeShapeIdentity {
  return !!identity
    && /^xl\/drawings\/(?!_rels\/)[^/]+\.xml$/.test(identity.drawingPart)
    && Number.isSafeInteger(identity.objectId)
    && identity.objectId > 0
    && identity.objectId <= 0xffff_ffff
}

function specId(identity: NativeShapeIdentity): string {
  return `native-shape-${identity.drawingPart.replace(/[^A-Za-z0-9_-]+/g, '_')}-${identity.objectId}`
}

/**
 * specFromFileShape converts one file shape into a mountable ShapeSpec.
 * Returns null when the kind isn't ours to render or its sheet is gone from
 * this workbook — the caller counts those rather than losing them silently.
 */
export function specFromFileShape(info: FileShapeInfo, sheetIdByName: Record<string, string>): FileShapeConversion | null {
  if (!RENDERABLE.has(info.kind)) return null
  if (!validIdentity(info.identity)) return null
  const sheetId = sheetIdByName[info.sheetName]
  if (!sheetId) return null
  const spec: ShapeSpec = {
    id: specId(info.identity),
    nativeIdentity: { ...info.identity },
    kind: info.kind as ShapeKind,
    text: info.text,
    fill: info.fill,
    stroke: info.stroke,
    strokeWidth: info.strokeWidth,
    textColor: info.textColor,
    fontSize: info.fontSize,
  }
  return { spec, cellAnchor: info.anchor, sheetId }
}

/** Convert a whole file's shapes; unrepresentable ones are counted, not
 *  lost silently. Grouped by sheet since the host mounts one sheet's
 *  shapes onto that sheet's active view. */
export function specsFromFileShapes(
  shapes: FileShapeInfo[],
  sheetIdByName: Record<string, string>,
): { conversions: FileShapeConversion[]; skipped: number } {
  const conversions: FileShapeConversion[] = []
  let skipped = 0
  const identities = new Set<string>()
  for (const s of shapes) {
    const identityKey = validIdentity(s.identity) ? `${s.identity.drawingPart}\u0000${s.identity.objectId}` : ''
    if (!identityKey || identities.has(identityKey)) {
      skipped++
      continue
    }
    const conv = specFromFileShape(s, sheetIdByName)
    if (conv) {
      identities.add(identityKey)
      conversions.push(conv)
    }
    else skipped++
  }
  return { conversions, skipped }
}
