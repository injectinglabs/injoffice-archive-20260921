import type { RenderPathCommand, RenderStroke } from './types.js'

/**
 * DrawingML ST_CompoundLine (ECMA-376 Part 1 §20.1.10.14) describes an outline
 * painted as several concentric lines separated by gaps, not as one line. The
 * `w` on `a:ln` is the width of the whole assembly, so a compound outline is
 * decomposed into bands that share that one width.
 *
 * Each entry below is the band list from the OUTER edge inward, as integer
 * weights of the total width. `paint` marks the bands carrying the outline
 * color; the others are gaps that reveal whatever the outline sits on — the
 * page outside the shape and the shape's own fill inside it.
 *
 * The weights are read off PowerPoint's own raster rather than assumed. On
 * hard-v2 `ShapeLineProperties.pptx` a `thinThick` outline of w=63500 EMU on a
 * 9144000 EMU page rasterizes at 960px to 6.667px total, and PowerPoint paints
 * the outline color at 217.33..219.00 px, white at 219.00..220.67, and the
 * outline color again at 220.67..224.00, against a shape path whose left edge
 * is 220.67. That is 1:1:2 over four parts, centered on the path (`algn="ctr"`
 * places w/2 outside and w/2 inside). The other members are the same rule with
 * the weight vector their name states.
 */
const COMPOUND_BANDS = {
  single: [{ weight: 1, paint: true }],
  double: [{ weight: 1, paint: true }, { weight: 1, paint: false }, { weight: 1, paint: true }],
  thickThin: [{ weight: 2, paint: true }, { weight: 1, paint: false }, { weight: 1, paint: true }],
  thinThick: [{ weight: 1, paint: true }, { weight: 1, paint: false }, { weight: 2, paint: true }],
  triple: [
    { weight: 1, paint: true }, { weight: 1, paint: false }, { weight: 1, paint: true },
    { weight: 1, paint: false }, { weight: 1, paint: true },
  ],
} as const satisfies Record<NonNullable<RenderStroke['compound']>, readonly { readonly weight: number; readonly paint: boolean }[]>

/** One painted band: its center offset inward from the shape path, and its width. */
export interface CompoundStrokeBand { readonly offsetEmu: number; readonly widthEmu: number }

/**
 * The painted bands of a compound outline, outer edge inward. Returns undefined
 * for the single-line default, which every existing caller already paints as
 * one centered stroke of the full width.
 */
export function compoundStrokeBands(stroke: RenderStroke): readonly CompoundStrokeBand[] | undefined {
  const kind = stroke.compound
  if (kind === undefined || kind === 'single') return undefined
  const bands = COMPOUND_BANDS[kind]
  const total = bands.reduce((sum, band) => sum + band.weight, 0)
  // Round each band edge once, so adjacent bands share an edge exactly and the
  // assembly still spans the authored width. Edges are measured inward from the
  // path, the outermost sitting at -w/2. A band of odd width has no integer
  // center, so its offset rounds by half an EMU — 1/1828800 of an inch, three
  // orders of magnitude below one device pixel at any print resolution.
  const outerEdge = -Math.round(stroke.widthEmu / 2)
  const edgeAt = (cursor: number) => outerEdge + Math.round(stroke.widthEmu * cursor / total)
  const painted: CompoundStrokeBand[] = []
  let cursor = 0
  for (const band of bands) {
    const from = edgeAt(cursor)
    cursor += band.weight
    const to = edgeAt(cursor)
    if (band.paint && to > from) painted.push({ offsetEmu: Math.round((from + to) / 2) + 0, widthEmu: to - from })
  }
  return painted.length ? painted : undefined
}

/**
 * The same axis-aligned rectangle moved `offsetEmu` inward on every side, with
 * its winding and starting corner preserved. Returns undefined when the path is
 * not an axis-aligned rectangle or the inset would collapse it, so the caller
 * can fall back to the single centered stroke rather than invent geometry.
 *
 * Only rectangles are offset here: insetting an arbitrary preset outline is a
 * different problem (joins move along their miter, not along the edge normal),
 * and the extractor refuses a compound outline on anything but a rectangle.
 */
export function offsetRectanglePath(path: readonly RenderPathCommand[], offsetEmu: number): readonly RenderPathCommand[] | undefined {
  // The rect preset compiles to the rectangle primitive, not to four corners.
  const primitive = path.length === 1 ? path[0] : undefined
  if (primitive?.kind === 'rect') {
    const { x, y, cx, cy } = primitive.rect
    if (cx <= 2 * offsetEmu || cy <= 2 * offsetEmu) return undefined
    return [{ kind: 'rect', rect: { x: x + offsetEmu, y: y + offsetEmu, cx: cx - 2 * offsetEmu, cy: cy - 2 * offsetEmu } }]
  }
  if (path.length !== 5 || path[0]?.kind !== 'moveTo' || path[4]?.kind !== 'close') return undefined
  const corners: { x: number; y: number }[] = []
  for (const command of path.slice(0, 4)) {
    if (command.kind !== 'moveTo' && command.kind !== 'lineTo') return undefined
    corners.push({ x: command.x, y: command.y })
  }
  const xs = [...new Set(corners.map((corner) => corner.x))], ys = [...new Set(corners.map((corner) => corner.y))]
  if (xs.length !== 2 || ys.length !== 2) return undefined
  const minX = Math.min(xs[0]!, xs[1]!), maxX = Math.max(xs[0]!, xs[1]!)
  const minY = Math.min(ys[0]!, ys[1]!), maxY = Math.max(ys[0]!, ys[1]!)
  // Consecutive corners of an axis-aligned rectangle always share exactly one
  // coordinate; four corners over two distinct x and two distinct y that also
  // alternate is the rectangle, and nothing else.
  for (let index = 0; index < 4; index += 1) {
    const here = corners[index]!, next = corners[(index + 1) % 4]!
    if ((here.x === next.x) === (here.y === next.y)) return undefined
  }
  if (maxX - minX <= 2 * offsetEmu || maxY - minY <= 2 * offsetEmu) return undefined
  const moved = corners.map((corner) => ({
    x: corner.x === minX ? minX + offsetEmu : maxX - offsetEmu,
    y: corner.y === minY ? minY + offsetEmu : maxY - offsetEmu,
  }))
  return [
    { kind: 'moveTo', x: moved[0]!.x, y: moved[0]!.y },
    ...moved.slice(1).map((corner) => ({ kind: 'lineTo' as const, x: corner.x, y: corner.y })),
    { kind: 'close' },
  ]
}
