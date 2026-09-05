import { useEffect, useReducer, useRef } from 'react'
import {
  arcPath,
  bowtiePath,
  bracePath,
  bracketPath,
  calloutPath,
  type CornerTreatment,
  cutCornerRectPath,
  diagStripePath,
  diamondPath,
  directionalArrowPath,
  displayShapePath,
  doubleArrowPath,
  dShapePath,
  ellipseCalloutPath,
  foldedCornerFlapPath,
  foldedCornerPath,
  framePath,
  homePlatePath,
  invertedTrianglePath,
  lBandPath,
  lineCalloutPath,
  moonPath,
  openArcPath,
  parallelogramPath,
  pillPath,
  plusPath,
  regularPolygonPath,
  rightTrianglePath,
  ringSegmentPath,
  roundRectPath,
  starPath,
  sunRaysPath,
  trapezoidPath,
  trianglePath,
  wavyRectPath,
} from './geometry'
import { getShapeSpec, pushShapeTextEdit, subscribeShape } from './registry'
import { hasText, shapeLabel, type ShapeKind } from './types'

// The React component Univer mounts per shape float-dom. Renders SVG geometry
// sized to the float (viewBox tracks the container via ResizeObserver) with a
// contentEditable text layer for text-bearing kinds. Text edits push back to
// the manager on blur — not per keystroke, so undo/typing stay smooth and the
// spec (and snapshot data) settle once per edit.
//
// Rendering fidelity, by design (S6.2/S6.3, 119 presets): the polygon/star/
// arrow/rect-corner/flowchart families are drawn from parametrized
// generators (geometry.ts) that cover dozens of kinds each — real, correct
// geometry, not one hand-authored path per preset. A small "novelty/long-
// tail" set (heart, cloud, ribbons, irregularSeal2, the magnetic-storage
// flowchart symbols...) gets a clean generic placeholder (dashed outline +
// the preset's name) here in the BROWSER PREVIEW ONLY — the SAVED FILE
// always carries the exact, correct OOXML prst regardless, and opens with
// its real look in Excel/PowerPoint/Google Sheets/LibreOffice, which is the
// actual deliverable. See docs/ROADMAP.md for exactly which kinds fall in
// which bucket.

const STAR_POINTS: Record<string, number> = {
  star4: 4, star5: 5, star6: 6, star7: 7, star8: 8,
  star10: 10, star12: 12, star16: 16, star24: 24, star32: 32,
}

const REGULAR_POLYGON_SIDES: Record<string, number> = {
  pentagon: 5, hexagon: 6, heptagon: 7, octagon: 8, decagon: 10, dodecagon: 12,
}

// S6.3: the 7 rect-corner presets, each a distinct corner-treatment config
// for the one cutCornerRectPath generator.
const CUT_CORNER_CONFIGS: Record<string, [CornerTreatment, CornerTreatment, CornerTreatment, CornerTreatment]> = {
  round1Rect: ['round', 'none', 'none', 'none'],
  round2SameRect: ['round', 'round', 'none', 'none'],
  round2DiagRect: ['round', 'none', 'round', 'none'],
  snip1Rect: ['none', 'snip', 'none', 'none'],
  snip2SameRect: ['snip', 'snip', 'none', 'none'],
  snip2DiagRect: ['snip', 'none', 'snip', 'none'],
  snipRoundRect: ['round', 'snip', 'round', 'snip'],
}

// The exotic arrow bends/curves this doesn't draw distinctly — approximated
// as a plain right-pointing block arrow (still visibly "an arrow", unlike a
// blank box) with the real preset name captioned in the corner.
const ARROW_APPROX: ReadonlySet<string> = new Set([
  'quadArrow', 'bentArrow', 'bentUpArrow', 'uturnArrow', 'leftUpArrow',
  'curvedRightArrow', 'curvedLeftArrow', 'curvedUpArrow', 'curvedDownArrow',
  'stripedRightArrow', 'notchedRightArrow', 'chevron', 'circularArrow',
])

const OTHER_APPROX: ReadonlySet<string> = new Set([
  'flowChartDocument', 'halfFrame', 'plaque', 'wedgeRoundRectCallout',
])

const GENERIC_PREVIEW: ReadonlySet<string> = new Set([
  'teardrop', 'cube', 'can', 'heart', 'lightningBolt', 'cloud', 'smileyFace',
  'cloudCallout', 'ribbon', 'ribbon2', 'ellipseRibbon', 'irregularSeal1',
  'irregularSeal2', 'flowChartManualInput', 'flowChartMagneticDisk',
  'flowChartMagneticDrum',
])

export type ShapePreviewFidelity = 'distinct' | 'approximated' | 'generic'

/** Describes browser-preview fidelity only; saved files always retain kind. */
export function shapePreviewFidelity(kind: ShapeKind): ShapePreviewFidelity {
  if (ARROW_APPROX.has(kind) || OTHER_APPROX.has(kind)) return 'approximated'
  if (GENERIC_PREVIEW.has(kind)) return 'generic'
  return 'distinct'
}

function FallbackBody({ w, h, stroke, kind }: { w: number; h: number; stroke: string; kind: string }) {
  const pad = 2
  return (
    <>
      <rect
        x={pad} y={pad} width={Math.max(0, w - 2 * pad)} height={Math.max(0, h - 2 * pad)}
        rx={6} fill="none" stroke={stroke === 'none' ? '#9aa5ab' : stroke} strokeWidth={1.25}
        strokeDasharray="4 3"
      />
      <text x={w / 2} y={12} textAnchor="middle" fontSize={9} fill="#7a848a">
        {shapeLabel(kind)}
      </text>
    </>
  )
}

function renderShapeSvg(kind: string, w: number, h: number, fill: string, stroke: string, sw: number): React.ReactNode {
  const pad = sw + 1
  if (kind === 'rect' || kind === 'flowChartProcess' || kind === 'flowChartDocument') {
    return <rect x={pad} y={pad} width={Math.max(0, w - 2 * pad)} height={Math.max(0, h - 2 * pad)} rx={6} fill={fill} stroke={stroke} strokeWidth={sw} />
  }
  if (kind === 'roundRect' || kind === 'flowChartAlternateProcess') {
    return <path d={roundRectPath(w, h)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  }
  if (kind === 'ellipse' || kind === 'flowChartConnector') {
    return <ellipse cx={w / 2} cy={h / 2} rx={Math.max(0, w / 2 - pad)} ry={Math.max(0, h / 2 - pad)} fill={fill} stroke={stroke} strokeWidth={sw} />
  }
  if (kind === 'flowChartTerminator') {
    return <path d={pillPath(w, h)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  }
  if (kind === 'triangle') return <path d={trianglePath(w, h)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  if (kind === 'rtTriangle') return <path d={rightTrianglePath(w, h)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  if (kind === 'diamond' || kind === 'flowChartDecision') return <path d={diamondPath(w, h)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  if (kind === 'parallelogram' || kind === 'flowChartInputOutput') return <path d={parallelogramPath(w, h)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  if (kind === 'trapezoid') return <path d={trapezoidPath(w, h)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  if (kind === 'plus') return <path d={plusPath(w, h)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  if (kind === 'flowChartPreparation') return <path d={regularPolygonPath(6, w, h)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  if (REGULAR_POLYGON_SIDES[kind]) {
    return <path d={regularPolygonPath(REGULAR_POLYGON_SIDES[kind], w, h)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  }
  if (STAR_POINTS[kind]) {
    return <path d={starPath(STAR_POINTS[kind], w, h)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  }
  if (kind === 'pie' || kind === 'chord' || kind === 'donut') {
    return <path d={arcPath(kind, w, h)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" fillRule="evenodd" />
  }
  if (kind === 'rightArrow' || kind === 'leftArrow' || kind === 'upArrow' || kind === 'downArrow') {
    return <path d={directionalArrowPath(kind.replace('Arrow', '').toLowerCase() as 'right' | 'left' | 'up' | 'down', w, h)} fill={stroke} stroke="none" strokeLinejoin="round" />
  }
  if (kind === 'leftRightArrow') return <path d={doubleArrowPath('horizontal', w, h)} fill="none" stroke={stroke} strokeWidth={sw} strokeLinecap="round" strokeLinejoin="round" />
  if (kind === 'upDownArrow') return <path d={doubleArrowPath('vertical', w, h)} fill="none" stroke={stroke} strokeWidth={sw} strokeLinecap="round" strokeLinejoin="round" />
  if (ARROW_APPROX.has(kind)) {
    return (
      <>
        <path d={directionalArrowPath('right', w, h)} fill={stroke} stroke="none" strokeLinejoin="round" />
        <text x={w - 3} y={h - 4} textAnchor="end" fontSize={8} fill="#7a848a">{shapeLabel(kind)}</text>
      </>
    )
  }
  if (kind === 'wedgeRectCallout' || kind === 'wedgeRoundRectCallout') {
    return <path d={calloutPath(w, h)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  }
  if (kind === 'wedgeEllipseCallout') {
    return <path d={ellipseCalloutPath(w, h)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  }
  if (kind === 'borderCallout1' || kind === 'borderCallout2' || kind === 'callout1' || kind === 'callout2') {
    return <path d={lineCalloutPath(w, h)} fill="none" stroke={stroke === 'none' ? '#59636a' : stroke} strokeWidth={sw} strokeLinecap="round" />
  }
  if (kind === 'flowChartSummingJunction') {
    const cx = w / 2, cy = h / 2, r = Math.min(w, h) / 2 - pad, k = r * 0.7
    return (
      <>
        <ellipse cx={cx} cy={cy} rx={r} ry={r} fill={fill} stroke={stroke} strokeWidth={sw} />
        <line x1={cx - k} y1={cy - k} x2={cx + k} y2={cy + k} stroke={stroke} strokeWidth={sw} />
        <line x1={cx - k} y1={cy + k} x2={cx + k} y2={cy - k} stroke={stroke} strokeWidth={sw} />
      </>
    )
  }
  if (kind === 'flowChartPredefinedProcess') {
    const barX = w * 0.12
    return (
      <>
        <rect x={pad} y={pad} width={Math.max(0, w - 2 * pad)} height={Math.max(0, h - 2 * pad)} fill={fill} stroke={stroke} strokeWidth={sw} />
        <line x1={barX} y1={pad} x2={barX} y2={h - pad} stroke={stroke} strokeWidth={sw} />
        <line x1={w - barX} y1={pad} x2={w - barX} y2={h - pad} stroke={stroke} strokeWidth={sw} />
      </>
    )
  }
  if (kind === 'flowChartInternalStorage') {
    const inset = Math.min(w, h) * 0.18
    return (
      <>
        <rect x={pad} y={pad} width={Math.max(0, w - 2 * pad)} height={Math.max(0, h - 2 * pad)} fill={fill} stroke={stroke} strokeWidth={sw} />
        <line x1={pad} y1={inset} x2={w - pad} y2={inset} stroke={stroke} strokeWidth={sw} />
        <line x1={inset} y1={pad} x2={inset} y2={h - pad} stroke={stroke} strokeWidth={sw} />
      </>
    )
  }

  // ---- S6.3: additional ECMA-376 presets ----
  if (CUT_CORNER_CONFIGS[kind]) {
    return <path d={cutCornerRectPath(w, h, CUT_CORNER_CONFIGS[kind])} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  }
  if (kind === 'mathPlus') {
    return <path d={plusPath(w, h, 0.14)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  }
  if (kind === 'frame') {
    return <path d={framePath(w, h)} fill={fill} stroke={stroke} strokeWidth={sw} fillRule="evenodd" />
  }
  if (kind === 'halfFrame' || kind === 'corner') {
    return <path d={lBandPath(w, h)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  }
  if (kind === 'diagStripe') {
    return <path d={diagStripePath(w, h)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  }
  if (kind === 'noSmoking') {
    const cx = w / 2, cy = h / 2, r = Math.min(w, h) / 2 - pad, k = r * 0.7
    return (
      <>
        <ellipse cx={cx} cy={cy} rx={r} ry={r} fill={fill} stroke={stroke} strokeWidth={sw} />
        <line x1={cx - k} y1={cy - k} x2={cx + k} y2={cy + k} stroke={stroke === 'none' ? '#59636a' : stroke} strokeWidth={sw} />
      </>
    )
  }
  if (kind === 'blockArc') {
    return <path d={ringSegmentPath(w, h, -45, 270)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  }
  if (kind === 'arc') {
    return <path d={openArcPath(w, h)} fill="none" stroke={stroke === 'none' ? '#59636a' : stroke} strokeWidth={sw} strokeLinecap="round" />
  }
  if (kind === 'foldedCorner') {
    return (
      <>
        <path d={foldedCornerPath(w, h)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
        <path d={foldedCornerFlapPath(w, h)} fill={stroke === 'none' ? '#c8d2ce' : stroke} stroke="none" opacity={0.55} />
      </>
    )
  }
  if (kind === 'bevel') {
    const inset = Math.min(w, h) * 0.16
    return (
      <>
        <rect x={pad} y={pad} width={Math.max(0, w - 2 * pad)} height={Math.max(0, h - 2 * pad)} fill={fill} stroke={stroke} strokeWidth={sw} />
        <rect x={inset} y={inset} width={Math.max(0, w - 2 * inset)} height={Math.max(0, h - 2 * inset)} fill="none" stroke={stroke === 'none' ? '#9aa5ab' : stroke} strokeWidth={Math.max(1, sw * 0.6)} opacity={0.6} />
      </>
    )
  }
  if (kind === 'sun') {
    const r = (Math.min(w, h) / 2) * 0.55
    return (
      <>
        <path d={sunRaysPath(w, h)} fill={stroke === 'none' ? fill : stroke} stroke="none" />
        <ellipse cx={w / 2} cy={h / 2} rx={r} ry={r} fill={fill} stroke={stroke} strokeWidth={sw} />
      </>
    )
  }
  if (kind === 'moon') {
    return <path d={moonPath(w, h)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  }
  if (kind === 'plaque') {
    return <path d={roundRectPath(w, h, 0.36)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  }
  if (kind === 'leftBracket' || kind === 'rightBracket') {
    return <path d={bracketPath(w, h, kind === 'leftBracket' ? 'left' : 'right')} fill="none" stroke={stroke === 'none' ? '#59636a' : stroke} strokeWidth={sw} strokeLinecap="round" />
  }
  if (kind === 'leftBrace' || kind === 'rightBrace') {
    return <path d={bracePath(w, h, kind === 'leftBrace' ? 'left' : 'right')} fill="none" stroke={stroke === 'none' ? '#59636a' : stroke} strokeWidth={sw} strokeLinecap="round" />
  }
  if (kind === 'homePlate') {
    return <path d={homePlatePath(w, h, 'right')} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  }
  if (kind === 'flowChartManualOperation') {
    return <path d={trapezoidPath(w, h, true)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  }
  if (kind === 'flowChartOffpageConnector') {
    return <path d={homePlatePath(w, h, 'down')} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  }
  if (kind === 'flowChartDisplay') {
    return <path d={displayShapePath(w, h)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  }
  if (kind === 'flowChartDelay') {
    return <path d={dShapePath(w, h, 'right')} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  }
  if (kind === 'flowChartOr') {
    const cx = w / 2, cy = h / 2, r = Math.min(w, h) / 2 - pad
    return (
      <>
        <ellipse cx={cx} cy={cy} rx={r} ry={r} fill={fill} stroke={stroke} strokeWidth={sw} />
        <line x1={cx} y1={cy - r} x2={cx} y2={cy + r} stroke={stroke} strokeWidth={sw} />
        <line x1={cx - r} y1={cy} x2={cx + r} y2={cy} stroke={stroke} strokeWidth={sw} />
      </>
    )
  }
  if (kind === 'flowChartCollate') {
    return <path d={bowtiePath(w, h)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  }
  if (kind === 'flowChartSort') {
    return (
      <>
        <path d={diamondPath(w, h)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
        <line x1={0} y1={h / 2} x2={w} y2={h / 2} stroke={stroke} strokeWidth={sw} />
      </>
    )
  }
  if (kind === 'flowChartExtract') {
    return <path d={trianglePath(w, h)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  }
  if (kind === 'flowChartMerge') {
    return <path d={invertedTrianglePath(w, h)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  }
  if (kind === 'flowChartPunchedTape') {
    return <path d={wavyRectPath(w, h, 'both', 3, 0.06)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  }
  if (kind === 'flowChartMultidocument') {
    const step = Math.min(w, h) * 0.08
    const bodyW = Math.max(0, w - 2 * step)
    const bodyH = Math.max(0, h - 2 * step)
    const layer = wavyRectPath(bodyW, bodyH, 'bottom', 2, 0.1)
    return (
      <>
        <path d={layer} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" transform="translate(0,0)" />
        <path d={layer} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" transform={`translate(${step},${step})`} />
        <path d={layer} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" transform={`translate(${2 * step},${2 * step})`} />
      </>
    )
  }
  // wave/doubleWave upgraded from the FallbackBody placeholder to real
  // geometry alongside the new flowChartPunchedTape, which needed the same
  // wavyRectPath generator anyway.
  if (kind === 'wave') {
    return <path d={wavyRectPath(w, h, 'bottom', 2, 0.12)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  }
  if (kind === 'doubleWave') {
    return <path d={wavyRectPath(w, h, 'both', 2, 0.1)} fill={fill} stroke={stroke} strokeWidth={sw} strokeLinejoin="round" />
  }

  return <FallbackBody w={w} h={h} stroke={stroke} kind={kind} />
}

function renderPreviewBody(kind: ShapeKind, w: number, h: number, fill: string, stroke: string, sw: number) {
  if (kind === 'line') {
    return <line x1={0} y1={h / 2} x2={w} y2={h / 2} stroke={stroke} strokeWidth={sw} strokeLinecap="round" />
  }
  if (kind === 'text') {
    return <text x={w / 2} y={h / 2} textAnchor="middle" dominantBaseline="middle" fill={stroke === 'none' ? '#59636a' : stroke} fontSize={Math.min(18, h * 0.24)}>Text</text>
  }
  return renderShapeSvg(kind, w, h, fill, stroke, sw)
}

export interface ShapePreviewProps {
  kind: ShapeKind
  width?: number
  height?: number
  fill?: string
  stroke?: string
  strokeWidth?: number
  className?: string
  ariaLabel?: string
  decorative?: boolean
}

/** Standalone SVG preview backed by the same dispatcher used by ShapeFloat. */
export function ShapePreview({
  kind,
  width = 180,
  height = 100,
  fill = 'none',
  stroke = 'none',
  strokeWidth = 1.5,
  className,
  ariaLabel,
  decorative = false,
}: ShapePreviewProps) {
  return (
    <svg
      className={className}
      viewBox={`0 0 ${width} ${height}`}
      role={decorative ? undefined : 'img'}
      aria-hidden={decorative || undefined}
      aria-label={decorative ? undefined : (ariaLabel ?? `${shapeLabel(kind)} browser preview`)}
      focusable="false"
    >
      {renderPreviewBody(kind, width, height, fill, stroke, strokeWidth)}
    </svg>
  )
}

export function ShapeFloat({ data }: { data?: { shapeId?: string } }) {
  const shapeId = data?.shapeId
  const hostRef = useRef<HTMLDivElement | null>(null)
  const [, force] = useReducer((x: number) => x + 1, 0)
  const sizeRef = useRef({ w: 200, h: 120 })

  useEffect(() => {
    if (!shapeId || !hostRef.current) return
    const unsubscribe = subscribeShape(shapeId, force)
    const ro = new ResizeObserver(() => {
      const el = hostRef.current
      if (!el || el.clientWidth === 0) return
      sizeRef.current = { w: el.clientWidth, h: el.clientHeight }
      force()
    })
    ro.observe(hostRef.current)
    return () => {
      unsubscribe()
      ro.disconnect()
    }
  }, [shapeId])

  const spec = shapeId ? getShapeSpec(shapeId) : undefined
  if (!spec) return <div ref={hostRef} style={{ width: '100%', height: '100%' }} />

  const { w, h } = sizeRef.current
  const fill = spec.fill || 'none'
  const stroke = spec.stroke || 'none'
  const sw = spec.strokeWidth ?? 1.5

  const svgBody = spec.kind === 'text' ? null : renderPreviewBody(spec.kind, w, h, fill, stroke, sw)

  const isCallout = spec.kind.startsWith('wedge') || spec.kind === 'cloudCallout'
  const textArea = hasText(spec.kind) ? (
    <div
      contentEditable
      suppressContentEditableWarning
      role="textbox"
      aria-label="Shape text"
      onBlur={(e) => pushShapeTextEdit(spec.id, e.currentTarget.textContent ?? '')}
      // Stop Univer's grid keyboard handling from stealing typing focus.
      onKeyDown={(e) => e.stopPropagation()}
      style={{
        position: 'absolute',
        inset: isCallout ? '6% 8% 26% 8%' : '8%',
        display: 'flex',
        alignItems: 'center',
        justifyContent: spec.kind === 'text' ? 'flex-start' : 'center',
        textAlign: spec.kind === 'text' ? 'left' : 'center',
        color: spec.textColor ?? '#1d2427',
        fontSize: spec.fontSize ?? 13,
        lineHeight: 1.35,
        outline: 'none',
        overflow: 'hidden',
        cursor: 'text',
      }}
    >
      {spec.text ?? ''}
    </div>
  ) : null

  return (
    <div ref={hostRef} style={{ position: 'relative', width: '100%', height: '100%' }}>
      <svg width="100%" height="100%" viewBox={`0 0 ${w} ${h}`} preserveAspectRatio="none" style={{ display: 'block' }}>
        {svgBody}
      </svg>
      {textArea}
    </div>
  )
}
