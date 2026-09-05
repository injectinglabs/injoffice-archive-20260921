// Pure SVG geometry for shape rendering — separated so path math is
// unit-testable independent of React/Univer. Parametrized generators
// (regular polygon, n-point star, directional arrow) each cover MANY of the
// 80 curated OOXML presets rather than one bespoke path per shape — real
// shapes fall out of a handful of formulas the same way OOXML's own preset
// geometry guides are parametrized (adjustment handles over a shared
// formula), not one hand-authored path per prst.

/** Regular n-gon inscribed in the w×h box, point-up. Covers pentagon(5)
 *  through dodecagon(12) — and triangle(3)/diamond(4) as callers that want
 *  the generic case rather than the hand-tuned isoceles/right variants. */
export function regularPolygonPath(sides: number, w: number, h: number): string {
  const cx = w / 2
  const cy = h / 2
  const rx = w / 2
  const ry = h / 2
  const points: string[] = []
  for (let i = 0; i < sides; i++) {
    // Point-up: start angle -90°, i.e. straight up, then step around.
    const angle = -Math.PI / 2 + (i * 2 * Math.PI) / sides
    points.push(`${cx + rx * Math.cos(angle)},${cy + ry * Math.sin(angle)}`)
  }
  return `M ${points.join(' L ')} Z`
}

/** n-pointed star inscribed in the w×h box. innerRatio controls point
 *  sharpness (OOXML's own star presets vary this per point-count; a flat
 *  ~0.42 reads well across 4–32 points without per-preset tuning). */
export function starPath(points: number, w: number, h: number, innerRatio = 0.42): string {
  const cx = w / 2
  const cy = h / 2
  const outerRx = w / 2
  const outerRy = h / 2
  const vertices: string[] = []
  const total = points * 2
  for (let i = 0; i < total; i++) {
    const angle = -Math.PI / 2 + (i * Math.PI) / points
    const isOuter = i % 2 === 0
    const rx = isOuter ? outerRx : outerRx * innerRatio
    const ry = isOuter ? outerRy : outerRy * innerRatio
    vertices.push(`${cx + rx * Math.cos(angle)},${cy + ry * Math.sin(angle)}`)
  }
  return `M ${vertices.join(' L ')} Z`
}

export type ArrowDirection = 'right' | 'left' | 'up' | 'down'

/** Single-direction block arrow: shaft + solid head, filling the w×h box.
 *  Generalizes the original right-pointing arrowPath to all 4 directions —
 *  covers rightArrow/leftArrow/upArrow/downArrow, and stands in as the
 *  "arrow-shaped" approximation for exotic bent/curved arrow presets that
 *  don't have (or aren't worth) real bespoke bend geometry. */
export function directionalArrowPath(dir: ArrowDirection, w: number, h: number): string {
  if (dir === 'up' || dir === 'down') {
    const midX = w / 2
    const headLen = Math.min(h * 0.28, w * 0.9, 26)
    const headHalf = Math.min(w / 2, headLen * 0.6)
    if (dir === 'up') {
      const shaftStart = Math.min(h, headLen)
      return `M ${midX} ${h} L ${midX} ${shaftStart} M ${midX - headHalf} ${shaftStart} L ${midX} 0 L ${midX + headHalf} ${shaftStart} Z`
    }
    const shaftEnd = Math.max(0, h - headLen)
    return `M ${midX} 0 L ${midX} ${shaftEnd} M ${midX - headHalf} ${shaftEnd} L ${midX} ${h} L ${midX + headHalf} ${shaftEnd} Z`
  }
  const midY = h / 2
  const headLen = Math.min(w * 0.28, h * 0.9, 26)
  const headHalf = Math.min(h / 2, headLen * 0.6)
  if (dir === 'left') {
    const shaftStart = Math.min(w, headLen)
    return `M ${w} ${midY} L ${shaftStart} ${midY} M ${shaftStart} ${midY - headHalf} L 0 ${midY} L ${shaftStart} ${midY + headHalf} Z`
  }
  const shaftEnd = Math.max(0, w - headLen)
  return `M 0 ${midY} L ${shaftEnd} ${midY} M ${shaftEnd} ${midY - headHalf} L ${w} ${midY} L ${shaftEnd} ${midY + headHalf} Z`
}

/** Two-headed arrow (leftRightArrow / upDownArrow): a head at both ends of
 *  the shaft, same head geometry as directionalArrowPath. */
export function doubleArrowPath(axis: 'horizontal' | 'vertical', w: number, h: number): string {
  if (axis === 'vertical') {
    const midX = w / 2
    const headLen = Math.min(h * 0.28, w * 0.9, 26)
    const headHalf = Math.min(w / 2, headLen * 0.6)
    return (
      `M ${midX} 0 L ${midX - headHalf} ${headLen} M ${midX} 0 L ${midX + headHalf} ${headLen} ` +
      `M ${midX} ${h} L ${midX - headHalf} ${h - headLen} M ${midX} ${h} L ${midX + headHalf} ${h - headLen} ` +
      `M ${midX} ${headLen} L ${midX} ${h - headLen}`
    )
  }
  const midY = h / 2
  const headLen = Math.min(w * 0.28, h * 0.9, 26)
  const headHalf = Math.min(h / 2, headLen * 0.6)
  return (
    `M 0 ${midY} L ${headLen} ${midY - headHalf} M 0 ${midY} L ${headLen} ${midY + headHalf} ` +
    `M ${w} ${midY} L ${w - headLen} ${midY - headHalf} M ${w} ${midY} L ${w - headLen} ${midY + headHalf} ` +
    `M ${headLen} ${midY} L ${w - headLen} ${midY}`
  )
}

/** Isoceles triangle, point-up. */
export function trianglePath(w: number, h: number): string {
  return `M ${w / 2} 0 L ${w} ${h} L 0 ${h} Z`
}

/** Isoceles triangle, point-down — flowChartMerge is trianglePath's mirror,
 *  the same relationship rightArrow/leftArrow have via directionalArrowPath. */
export function invertedTrianglePath(w: number, h: number): string {
  return `M 0 0 L ${w} 0 L ${w / 2} ${h} Z`
}

/** Right triangle, right angle at bottom-left. */
export function rightTrianglePath(w: number, h: number): string {
  return `M 0 0 L 0 ${h} L ${w} ${h} Z`
}

/** Diamond (rhombus): the 4 box midpoints. */
export function diamondPath(w: number, h: number): string {
  return `M ${w / 2} 0 L ${w} ${h / 2} L ${w / 2} ${h} L 0 ${h / 2} Z`
}

/** Parallelogram, slanted right. */
export function parallelogramPath(w: number, h: number): string {
  const slant = w * 0.2
  return `M ${slant} 0 L ${w} 0 L ${w - slant} ${h} L 0 ${h} Z`
}

/** Trapezoid; narrower on top by default, invert for flowChartManualOperation
 *  (wider on top, narrower on bottom — the same shape upside down). */
export function trapezoidPath(w: number, h: number, invert = false): string {
  const inset = w * 0.18
  if (invert) {
    return `M 0 0 L ${w} 0 L ${w - inset} ${h} L ${inset} ${h} Z`
  }
  return `M ${inset} 0 L ${w - inset} 0 L ${w} ${h} L 0 ${h} Z`
}

/** Plus / cross. armRatio controls arm thickness relative to the box's
 *  shorter side — the default (1/3) is the thick "plus" preset; mathPlus
 *  reuses this with a much thinner ratio for the slim "+" operator glyph,
 *  one formula covering both rather than a second hand-authored path.
 *  12 corners, explicit L points throughout (not H/V shorthand) — every
 *  other function in this file emits plain coordinate pairs, and
 *  consistency here keeps path output uniformly parseable. */
export function plusPath(w: number, h: number, armRatio = 1 / 3): string {
  const armW = Math.min(w, h) * armRatio
  const x0 = (w - armW) / 2
  const x1 = x0 + armW
  const y0 = (h - armW) / 2
  const y1 = y0 + armW
  const pts: Array<[number, number]> = [
    [x0, 0], [x1, 0], [x1, y0], [w, y0], [w, y1], [x1, y1],
    [x1, h], [x0, h], [x0, y1], [0, y1], [0, y0], [x0, y0],
  ]
  return `M ${pts.map(([x, y]) => `${x} ${y}`).join(' L ')} Z`
}

/** Rounded rectangle; radiusRatio of the shorter side. */
export function roundRectPath(w: number, h: number, radiusRatio = 0.16): string {
  const r = Math.min(w, h) * radiusRatio
  return (
    `M ${r} 0 H ${w - r} Q ${w} 0 ${w} ${r} V ${h - r} Q ${w} ${h} ${w - r} ${h} ` +
    `H ${r} Q 0 ${h} 0 ${h - r} V ${r} Q 0 0 ${r} 0 Z`
  )
}

/** Stadium / pill: fully rounded ends (flowchart terminator). */
export function pillPath(w: number, h: number): string {
  return roundRectPath(w, h, 0.5)
}

export type ArcMode = 'pie' | 'chord' | 'donut'

/** Pie / chord / donut: all the same circular-sector math, differing only
 *  in how the sector closes. A quarter-circle wedge (0°→90°) reads clearly
 *  at small sizes, which is what these annotate as (a data callout, not a
 *  literal percentage) — an exact angle isn't the point. */
export function arcPath(mode: ArcMode, w: number, h: number): string {
  const cx = w / 2
  const cy = h / 2
  const rx = w / 2
  const ry = h / 2
  const start = { x: cx, y: cy - ry }
  const end = { x: cx + rx, y: cy }
  if (mode === 'donut') {
    const innerRx = rx * 0.5
    const innerRy = ry * 0.5
    return (
      `M ${cx} ${cy - ry} A ${rx} ${ry} 0 1 1 ${cx - 0.01} ${cy - ry} Z ` +
      `M ${cx} ${cy - innerRy} A ${innerRx} ${innerRy} 0 1 0 ${cx - 0.01} ${cy - innerRy} Z`
    )
  }
  if (mode === 'chord') {
    return `M ${start.x} ${start.y} A ${rx} ${ry} 0 0 1 ${end.x} ${end.y} Z`
  }
  return `M ${cx} ${cy} L ${start.x} ${start.y} A ${rx} ${ry} 0 0 1 ${end.x} ${end.y} Z`
}

/** Callout body + tail, tail dropping from the bottom-left third. Rounded
 *  corners — reused for both wedgeRectCallout and wedgeRoundRectCallout
 *  (the visual difference is minor at annotation scale). */
export function calloutPath(w: number, h: number): string {
  const bodyH = h * 0.78
  const r = Math.min(8, w / 6, bodyH / 6)
  const tailX = w * 0.22
  return (
    `M ${r} 0 H ${w - r} Q ${w} 0 ${w} ${r} V ${bodyH - r} Q ${w} ${bodyH} ${w - r} ${bodyH} ` +
    `H ${tailX + 14} L ${tailX} ${h} L ${tailX + 4} ${bodyH} H ${r} Q 0 ${bodyH} 0 ${bodyH - r} ` +
    `V ${r} Q 0 0 ${r} 0 Z`
  )
}

/** Oval callout: an ellipse body with the same drop-tail. */
export function ellipseCalloutPath(w: number, h: number): string {
  const bodyH = h * 0.78
  const cx = w / 2
  const tailX = w * 0.3
  return (
    `M ${cx} 0 A ${w / 2} ${bodyH / 2} 0 1 1 ${cx - 0.01} 0 Z ` +
    `M ${tailX} ${bodyH * 0.85} L ${tailX - 6} ${h} L ${tailX + 18} ${bodyH * 0.9}`
  )
}

/** Line/border callout (borderCallout1/2, callout1/2): a text box with a
 *  bent leader line to a point outside it — approximates all four
 *  variants (which really only differ in bend count / border-or-not). */
export function lineCalloutPath(w: number, h: number): string {
  const boxY = h * 0.32
  return `M 0 ${boxY} L ${w * 0.18} ${boxY * 0.4} L ${w * 0.42} ${boxY}`
}

// ---- S6.3: additional ECMA-376 presets beyond the S6.2 catalogue ----
// Same philosophy as everything above: a handful of parametrized formulas
// covering many presets, not one hand-authored path per prst. A few of the
// 39 (irregularSeal2, the two magnetic-storage flowchart symbols) get the
// honest labeled fallback instead — the same treatment "can"/"cube" already
// get — rather than a fabricated-looking approximation.

export type CornerTreatment = 'none' | 'round' | 'snip'

/** Rectangle with independently treated corners (sharp / rounded / 45°-
 *  snipped) — covers round1Rect, round2SameRect, round2DiagRect, snip1Rect,
 *  snip2SameRect, snip2DiagRect and snipRoundRect from one formula, the same
 *  way roundRectPath already covers plain roundRect. corners = [topLeft,
 *  topRight, bottomRight, bottomLeft]. cutRatio sizes the treatment relative
 *  to the shorter side — a consistent, correct-looking cut/round at
 *  annotation scale, not a claim of matching each preset's exact ECMA-376
 *  default adj value. */
export function cutCornerRectPath(
  w: number,
  h: number,
  corners: [CornerTreatment, CornerTreatment, CornerTreatment, CornerTreatment],
  cutRatio = 0.18,
): string {
  const r = Math.min(w, h) * cutRatio
  const [tl, tr, br, bl] = corners
  const segs: string[] = []
  segs.push(`M ${tl === 'none' ? 0 : r} 0`)
  segs.push(`L ${tr === 'none' ? w : w - r} 0`)
  if (tr === 'round') segs.push(`Q ${w} 0 ${w} ${r}`)
  else if (tr === 'snip') segs.push(`L ${w} ${r}`)
  segs.push(`L ${w} ${br === 'none' ? h : h - r}`)
  if (br === 'round') segs.push(`Q ${w} ${h} ${w - r} ${h}`)
  else if (br === 'snip') segs.push(`L ${w - r} ${h}`)
  segs.push(`L ${bl === 'none' ? 0 : r} ${h}`)
  if (bl === 'round') segs.push(`Q 0 ${h} 0 ${h - r}`)
  else if (bl === 'snip') segs.push(`L 0 ${h - r}`)
  segs.push(`L 0 ${tl === 'none' ? 0 : r}`)
  if (tl === 'round') segs.push(`Q 0 0 ${r} 0`)
  else if (tl === 'snip') segs.push(`L ${r} 0`)
  segs.push('Z')
  return segs.join(' ')
}

/** Picture frame: an outer rect with a smaller rect punched out of the
 *  middle. Render with fillRule="evenodd" so the hole reads as empty
 *  regardless of the shape's own fill color. */
export function framePath(w: number, h: number, thicknessRatio = 0.16): string {
  const t = Math.min(w, h) * thicknessRatio
  const outer = `M 0 0 L ${w} 0 L ${w} ${h} L 0 ${h} Z`
  const inner = `M ${t} ${t} L ${t} ${h - t} L ${w - t} ${h - t} L ${w - t} ${t} Z`
  return `${outer} ${inner}`
}

/** L-shaped band along the top and left edges. "corner" (a plain L bracket)
 *  and "halfFrame" (an L-shaped picture frame) share this geometry — the
 *  visual difference between the two presets is minor at annotation scale,
 *  the same simplification calloutPath already makes for its two callers. */
export function lBandPath(w: number, h: number, thicknessRatio = 0.16): string {
  const t = Math.min(w, h) * thicknessRatio
  return `M 0 0 L ${w} 0 L ${w} ${t} L ${t} ${t} L ${t} ${h} L 0 ${h} Z`
}

/** Diagonal stripe/band across the box, corner to corner. */
export function diagStripePath(w: number, h: number): string {
  const band = Math.min(w, h) * 0.35
  return `M 0 0 L ${band} 0 L ${w} ${h - band} L ${w} ${h} L ${w - band} ${h} L 0 ${band} Z`
}

/** Thick partial ring (blockArc): an arc swept sweepDeg degrees starting at
 *  startDeg (0 = top, clockwise), with real thickness via innerRatio —
 *  unlike a plain stroked arc, this is a filled band. */
export function ringSegmentPath(
  w: number,
  h: number,
  startDeg: number,
  sweepDeg: number,
  innerRatio = 0.55,
): string {
  const cx = w / 2
  const cy = h / 2
  const rx = w / 2
  const ry = h / 2
  const irx = rx * innerRatio
  const iry = ry * innerRatio
  const toRad = (d: number) => (d * Math.PI) / 180
  const pt = (deg: number, radx: number, rady: number) => {
    const a = toRad(deg)
    return { x: cx + radx * Math.sin(a), y: cy - rady * Math.cos(a) }
  }
  const endDeg = startDeg + sweepDeg
  const large = sweepDeg > 180 ? 1 : 0
  const o1 = pt(startDeg, rx, ry)
  const o2 = pt(endDeg, rx, ry)
  const i2 = pt(endDeg, irx, iry)
  const i1 = pt(startDeg, irx, iry)
  return (
    `M ${o1.x} ${o1.y} A ${rx} ${ry} 0 ${large} 1 ${o2.x} ${o2.y} ` +
    `L ${i2.x} ${i2.y} A ${irx} ${iry} 0 ${large} 0 ${i1.x} ${i1.y} Z`
  )
}

/** Open elliptical arc (the "arc" preset — a curve, not a filled region;
 *  rendered with fill="none" regardless of the shape's own fill, same as
 *  the arrow family forces fill=stroke). startDeg/sweepDeg as
 *  ringSegmentPath (0 = top, clockwise). */
export function openArcPath(w: number, h: number, startDeg = -90, sweepDeg = 180): string {
  const cx = w / 2
  const cy = h / 2
  const rx = w / 2
  const ry = h / 2
  const toRad = (d: number) => (d * Math.PI) / 180
  const pt = (deg: number) => {
    const a = toRad(deg)
    return { x: cx + rx * Math.sin(a), y: cy - ry * Math.cos(a) }
  }
  const large = sweepDeg > 180 ? 1 : 0
  const p1 = pt(startDeg)
  const p2 = pt(startDeg + sweepDeg)
  return `M ${p1.x} ${p1.y} A ${rx} ${ry} 0 ${large} 1 ${p2.x} ${p2.y}`
}

/** Folded-corner card: the outer boundary with the bottom-right corner cut
 *  at 45°. The small triangular "flap" suggesting the fold is a SEPARATE
 *  path (foldedCornerFlapPath) so ShapeFloat can give it a distinct shade
 *  from the card body using colors the spec already carries (the stroke
 *  color), rather than inventing a third color field on ShapeSpec. */
export function foldedCornerPath(w: number, h: number, foldRatio = 0.22): string {
  const f = Math.min(w, h) * foldRatio
  return `M 0 0 L ${w} 0 L ${w} ${h - f} L ${w - f} ${h} L 0 ${h} Z`
}

export function foldedCornerFlapPath(w: number, h: number, foldRatio = 0.22): string {
  const f = Math.min(w, h) * foldRatio
  return `M ${w - f} ${h} L ${w - f} ${h - f} L ${w} ${h} Z`
}

/** Sun rays: thin triangles radiating from innerRatio out to the box edge.
 *  The circular body itself is a plain <ellipse> drawn separately in
 *  ShapeFloat — this returns only the ray path. */
export function sunRaysPath(w: number, h: number, rays = 8, innerRatio = 0.55): string {
  const cx = w / 2
  const cy = h / 2
  const outerR = Math.min(w, h) / 2
  const innerR = outerR * innerRatio
  const rayHalfAngle = (Math.PI / rays) * 0.35
  const segs: string[] = []
  for (let i = 0; i < rays; i++) {
    const mid = (i * 2 * Math.PI) / rays - Math.PI / 2
    const tipX = cx + outerR * Math.cos(mid)
    const tipY = cy + outerR * Math.sin(mid)
    const baseX1 = cx + innerR * Math.cos(mid - rayHalfAngle)
    const baseY1 = cy + innerR * Math.sin(mid - rayHalfAngle)
    const baseX2 = cx + innerR * Math.cos(mid + rayHalfAngle)
    const baseY2 = cy + innerR * Math.sin(mid + rayHalfAngle)
    segs.push(`M ${baseX1} ${baseY1} L ${tipX} ${tipY} L ${baseX2} ${baseY2} Z`)
  }
  return segs.join(' ')
}

/** Crescent moon: an outer half-circle (radius rx) and an inner half-circle
 *  (smaller radius, opposite sweep) sharing the same top/bottom endpoints —
 *  the standard two-arc construction for a lune. */
export function moonPath(w: number, h: number): string {
  const cx = w / 2
  const cy = h / 2
  const rx = w / 2
  const ry = h / 2
  const innerRx = rx * 0.62
  return (
    `M ${cx} ${cy - ry} A ${rx} ${ry} 0 1 0 ${cx} ${cy + ry} ` +
    `A ${innerRx} ${ry} 0 1 1 ${cx} ${cy - ry} Z`
  )
}

/** Bracket glyph ("[" or "]"): a curved hook hugging one vertical edge — an
 *  annotation mark, not a filled region (render with fill="none"). */
export function bracketPath(w: number, h: number, side: 'left' | 'right'): string {
  const depth = Math.min(w, h) * 0.4
  const capH = h * 0.15
  const x = side === 'left' ? 0 : w
  const dx = side === 'left' ? depth : -depth
  return `M ${x + dx} 0 Q ${x} 0 ${x} ${capH} L ${x} ${h - capH} Q ${x} ${h} ${x + dx} ${h}`
}

/** Brace glyph ("{" or "}"): two curves meeting at a small inward-pointing
 *  tip at vertical center — an annotation mark, not a filled region. */
export function bracePath(w: number, h: number, side: 'left' | 'right'): string {
  const depth = Math.min(w, h) * 0.45
  const tipDepth = depth * 1.3
  const midY = h / 2
  const x = side === 'left' ? 0 : w
  const outward = side === 'left' ? depth : -depth
  const tipX = side === 'left' ? x + tipDepth : x - tipDepth
  return (
    `M ${x + outward} 0 Q ${x} 0 ${x} ${h * 0.2} Q ${x} ${midY * 0.85} ${tipX} ${midY} ` +
    `Q ${x} ${midY * 1.15} ${x} ${h * 0.8} Q ${x} ${h} ${x + outward} ${h}`
  )
}

/** Home-plate tab: a rectangle with one edge replaced by a point — rightward
 *  for homePlate (a "next step" tab shape), downward for the off-page-
 *  connector flowchart symbol. */
export function homePlatePath(w: number, h: number, dir: 'right' | 'down' = 'right'): string {
  if (dir === 'down') {
    const tip = h * 0.7
    return `M 0 0 L ${w} 0 L ${w} ${tip} L ${w / 2} ${h} L 0 ${tip} Z`
  }
  const tip = w * 0.75
  return `M 0 0 L ${tip} 0 L ${w} ${h / 2} L ${tip} ${h} L 0 ${h} Z`
}

/** "D" shape: flat on one side, semicircular on the other — flowChartDelay
 *  (flat left, rounded right by default). */
export function dShapePath(w: number, h: number, side: 'left' | 'right' = 'right'): string {
  const r = h / 2
  if (side === 'right') {
    return `M 0 0 L ${w - r} 0 A ${r} ${r} 0 0 1 ${w - r} ${h} L 0 ${h} Z`
  }
  return `M ${r} 0 L ${w} 0 L ${w} ${h} L ${r} ${h} A ${r} ${r} 0 0 1 ${r} 0 Z`
}

/** Bowtie / hourglass (flowChartCollate): two triangles meeting at a point,
 *  a single self-intersecting quadrilateral. */
export function bowtiePath(w: number, h: number): string {
  return `M 0 0 L ${w} 0 L 0 ${h} L ${w} ${h} Z`
}

/** CRT-display flowchart symbol: a concave point on the left (two straight
 *  edges meeting at the vertical center), a smooth rounded curve on the
 *  right. */
export function displayShapePath(w: number, h: number): string {
  const curveX = w * 0.72
  const notchX = w * 0.2
  return (
    `M ${notchX} 0 L ${curveX} 0 Q ${w} 0 ${w} ${h / 2} Q ${w} ${h} ${curveX} ${h} ` +
    `L ${notchX} ${h} L 0 ${h / 2} Z`
  )
}

/** One horizontal wavy edge, sampled as a polyline (not a handful of bezier
 *  bumps) — trivial to reverse for a right-to-left traversal and trivial to
 *  test against the exact sine formula. baseline y, waveCount full
 *  oscillations across width w, amplitude amp. */
function waveSamples(w: number, y: number, waveCount: number, amp: number, samplesPerWave = 8): Array<[number, number]> {
  const n = waveCount * samplesPerWave
  const pts: Array<[number, number]> = []
  for (let i = 0; i <= n; i++) {
    const x = (w * i) / n
    const yy = y + amp * Math.sin((i / samplesPerWave) * 2 * Math.PI)
    pts.push([x, yy])
  }
  return pts
}

function polyline(pts: Array<[number, number]>, continued: boolean): string {
  return pts.map(([x, y], i) => `${i === 0 && !continued ? 'M' : 'L'} ${x} ${y}`).join(' ')
}

/** Rectangle with one or both horizontal edges replaced by a gentle sine
 *  wave — covers wave (bottom wavy), doubleWave (both wavy, the bottom
 *  phase-inverted so the two edges don't just mirror into a uniform-
 *  thickness band), and flowChartPunchedTape (both wavy, tighter/shallower)
 *  from one formula. */
export function wavyRectPath(
  w: number,
  h: number,
  edges: 'bottom' | 'both' = 'bottom',
  waveCount = 2,
  amplitudeRatio = 0.08,
): string {
  const amp = h * amplitudeRatio
  if (edges === 'bottom') {
    const bottom = waveSamples(w, h, waveCount, amp).reverse()
    return `M 0 0 L ${w} 0 ${polyline(bottom, true)} Z`
  }
  const top = waveSamples(w, 0, waveCount, amp)
  const bottom = waveSamples(w, h, waveCount, -amp).reverse()
  return `${polyline(top, false)} ${polyline(bottom, true)} Z`
}
