import type { DeckTheme } from './types'
import type { DiagramSpec, OrgChartNode } from './types'
import type { WireShape } from './wire'

// S12 diagram slice: 'process' and 'orgChart' DiagramSpec -> WireShape[],
// laid out here (not stored pre-laid-out in the spec) and compiled to the
// SAME primitives every other compileXxx in compile.ts already produces
// (roundRect boxes + KindLine connectors) — see SlideSpec.diagram's doc
// comment in types.ts for why this is a deliberate stand-in for real OOXML
// SmartArt (the `dgm` namespace) rather than an attempt at it. Pure
// functions: same input, same output, no rendering — same contract as
// compile.ts.
//
// Scope, deliberately narrow (first slice — see the S12 round's report for
// what's NOT here): single-row process flow (no wrapping to multiple
// rows), and an org chart with no box-count/depth limit but no collision
// avoidance either — a very wide or very deep tree will produce narrow or
// vertically-overflowing boxes rather than dropping content or erroring.
// Both degrade gracefully (never data loss, never invalid output), just not
// beautifully, past a reasonable size.

const EMU_PER_INCH = 914400
function inch(n: number): number {
  return Math.round(n * EMU_PER_INCH)
}

export interface DiagramRegion {
  x: number
  y: number
  w: number
  h: number
}

function labelBox(x: number, y: number, w: number, h: number, text: string, theme: DeckTheme, key: string): WireShape {
  return {
    kind: 'roundRect',
    x: inch(x),
    y: inch(y),
    cx: inch(w),
    cy: inch(h),
    fill: theme.accent,
    stroke: theme.ink,
    strokeWidthPt: 0.75,
    paragraphs: [{ runs: [{ text, color: '#ffffff', sizePt: 13, bold: true }], align: 'ctr' }],
    key,
  }
}

/** Straight connector between two arbitrary points, normalized to a
 *  non-negative a:xfrm bounding box with FlipH set when the two points fall
 *  on the box's OTHER (anti-) diagonal. Mirrors go/pptxpatch's
 *  DiagramExampleDeck `connector` helper 1:1 — see Shape.FlipH's doc
 *  comment there for the full derivation. headArrow/tailArrow let process
 *  flow reuse this for its directional arrows (always non-diagonal there,
 *  so flip is always false, but sharing the one function keeps write/read
 *  symmetry obvious). */
function connector(
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  theme: DeckTheme,
  key: string,
  opts: { tailArrow?: boolean; headArrow?: boolean } = {},
): WireShape {
  const offX = Math.min(x1, x2)
  const offY = Math.min(y1, y2)
  const cx = Math.abs(x2 - x1)
  const cy = Math.abs(y2 - y1)
  const flipH = cx > 0 && cy > 0 && x1 < x2 !== y1 < y2
  const shape: WireShape = {
    kind: 'line',
    x: inch(offX),
    y: inch(offY),
    cx: inch(cx),
    cy: inch(cy),
    stroke: theme.muted,
    strokeWidthPt: opts.tailArrow || opts.headArrow ? 1.5 : 1,
    key,
  }
  if (flipH) shape.flipH = true
  if (opts.headArrow) shape.headArrow = true
  if (opts.tailArrow) shape.tailArrow = true
  return shape
}

// ---- process flow: a single left-to-right row of boxes + arrows ----

const PROCESS_GAP_IN = 0.35
const PROCESS_MIN_BOX_W_IN = 0.5

function compileProcessFlow(steps: string[], theme: DeckTheme, region: DiagramRegion): WireShape[] {
  const n = steps.length
  if (n === 0) return []
  const boxH = Math.min(1.1, region.h)
  const y = region.y + (region.h - boxH) / 2
  const totalGap = PROCESS_GAP_IN * (n - 1)
  const boxW = Math.max(PROCESS_MIN_BOX_W_IN, (region.w - totalGap) / n)

  const shapes: WireShape[] = []
  let x = region.x
  const centers: number[] = []
  steps.forEach((label, i) => {
    shapes.push(labelBox(x, y, boxW, boxH, label, theme, `diagram-step-${i}`))
    centers.push(x)
    x += boxW + PROCESS_GAP_IN
  })
  for (let i = 0; i < n - 1; i++) {
    const x1 = centers[i]! + boxW
    const x2 = centers[i + 1]!
    shapes.push(connector(x1, y + boxH / 2, x2, y + boxH / 2, theme, `diagram-arrow-${i}`, { tailArrow: true }))
  }
  return shapes
}

// ---- org chart: a top-down tree, parent bottom-center -> child top-center ----

const ORGCHART_BOX_H_IN = 0.7
const ORGCHART_ROW_GAP_IN = 0.55
const ORGCHART_MIN_BOX_W_IN = 0.9
const ORGCHART_MAX_BOX_W_IN = 1.9

interface PositionedNode {
  node: OrgChartNode
  x: number
  y: number
  cx: number
  cy: number
  children: PositionedNode[]
}

function countLeaves(n: OrgChartNode): number {
  if (!n.children || n.children.length === 0) return 1
  return n.children.reduce((sum, c) => sum + countLeaves(c), 0)
}

function layoutOrgChart(root: OrgChartNode, region: DiagramRegion): PositionedNode {
  const totalLeaves = Math.max(1, countLeaves(root))
  const leafSlotW = region.w / totalLeaves
  const boxW = Math.max(ORGCHART_MIN_BOX_W_IN, Math.min(ORGCHART_MAX_BOX_W_IN, leafSlotW * 0.78))

  let nextSlot = 0
  function place(node: OrgChartNode, depth: number): PositionedNode {
    const children = (node.children ?? []).map((c) => place(c, depth + 1))
    let centerSlot: number
    if (children.length === 0) {
      centerSlot = nextSlot + 0.5
      nextSlot += 1
    } else {
      centerSlot = children.reduce((sum, c) => sum + (c.x + c.cx / 2 - region.x) / leafSlotW, 0) / children.length
    }
    return {
      node,
      x: region.x + centerSlot * leafSlotW - boxW / 2,
      y: region.y + depth * (ORGCHART_BOX_H_IN + ORGCHART_ROW_GAP_IN),
      cx: boxW,
      cy: ORGCHART_BOX_H_IN,
      children,
    }
  }
  return place(root, 0)
}

function flattenOrgChart(pn: PositionedNode, theme: DeckTheme, shapes: WireShape[], counter: { n: number }): void {
  const myIndex = counter.n++
  shapes.push(labelBox(pn.x, pn.y, pn.cx, pn.cy, pn.node.label, theme, `diagram-org-${myIndex}`))
  for (const child of pn.children) {
    const x1 = pn.x + pn.cx / 2
    const y1 = pn.y + pn.cy
    const x2 = child.x + child.cx / 2
    const y2 = child.y
    shapes.push(connector(x1, y1, x2, y2, theme, `diagram-org-edge-${myIndex}-${counter.n}`))
    flattenOrgChart(child, theme, shapes, counter)
  }
}

function compileOrgChart(root: OrgChartNode, theme: DeckTheme, region: DiagramRegion): WireShape[] {
  const positioned = layoutOrgChart(root, region)
  const shapes: WireShape[] = []
  flattenOrgChart(positioned, theme, shapes, { n: 0 })
  return shapes
}

/** Compile a DiagramSpec into WireShapes positioned within `region` (inches).
 *  undefined diagram -> no shapes (the common case: most slides carry none). */
export function compileDiagramShapes(diagram: DiagramSpec | undefined, theme: DeckTheme, region: DiagramRegion): WireShape[] {
  if (!diagram) return []
  if (diagram.type === 'process') return compileProcessFlow(diagram.steps, theme, region)
  return compileOrgChart(diagram.root, theme, region)
}
