// Shape model for @injoffice/shapes.
//
// Same contract philosophy as ChartSpec/PivotSpec: plain JSON, mountable on
// Univer's float-DOM layer (move/resize come free via allowTransform).
//
// A shape's KIND *is* its OOXML preset geometry name (a:prstGeom prst="...")
// directly, drawn from ECMA-376's own preset enumeration (182 total,
// verified against python-pptx's MSO_AUTO_SHAPE_TYPE) — no translation
// layer between "what we call it" and "what the file says". Two kinds are
// structural specials, not presets: "line" (an OOXML connector, not an
// auto-shape) and "text" (prst="rect" marked with the real txBox="1" idiom
// — see the Go writer, xlsxpatch/shapewrite.go).
//
// S6.2 shipped the curated ~80-shape subset — the ones people actually
// reach for, not the full 182. S6.3 (2026-08-25) added 39 more ECMA-376
// presets (7 rect-corner variants — round1Rect..snipRoundRect — pulled into
// their own "Rects" group so Basic stayed a usable toolbar group; 18 more
// basic-shape extras — frame, bevel, brackets/braces, sun/moon, homePlate...;
// 13 more flowchart symbols; 1 more star/banner). Grouped for the toolbar's
// <optgroup> picker.

export interface ShapeCategory {
  label: string
  kinds: readonly string[]
}

const BASIC = [
  'rect', 'roundRect', 'ellipse', 'triangle', 'rtTriangle', 'parallelogram', 'trapezoid',
  'diamond', 'pentagon', 'hexagon', 'heptagon', 'octagon', 'decagon', 'dodecagon', 'plus',
  'pie', 'chord', 'donut', 'teardrop', 'cube', 'can', 'heart', 'lightningBolt', 'cloud',
  'smileyFace',
  'mathPlus', 'frame', 'halfFrame', 'corner', 'diagStripe', 'noSmoking', 'blockArc',
  'foldedCorner', 'bevel', 'sun', 'moon', 'arc', 'plaque', 'leftBracket', 'rightBracket',
  'leftBrace', 'rightBrace', 'homePlate',
] as const

const RECTS = [
  'round1Rect', 'round2SameRect', 'round2DiagRect',
  'snip1Rect', 'snip2SameRect', 'snip2DiagRect', 'snipRoundRect',
] as const

const ARROWS = [
  'rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'leftRightArrow', 'upDownArrow',
  'quadArrow', 'bentArrow', 'bentUpArrow', 'uturnArrow', 'leftUpArrow', 'curvedRightArrow',
  'curvedLeftArrow', 'curvedUpArrow', 'curvedDownArrow', 'stripedRightArrow',
  'notchedRightArrow', 'chevron', 'circularArrow',
] as const

const CALLOUTS = [
  'wedgeRectCallout', 'wedgeRoundRectCallout', 'wedgeEllipseCallout', 'cloudCallout',
  'borderCallout1', 'borderCallout2', 'callout1', 'callout2',
] as const

const STARS_BANNERS = [
  'star4', 'star5', 'star6', 'star7', 'star8', 'star10', 'star12', 'star16', 'star24',
  'star32', 'ribbon', 'ribbon2', 'ellipseRibbon', 'wave', 'doubleWave', 'irregularSeal1',
  'irregularSeal2',
] as const

const FLOWCHART = [
  'flowChartProcess', 'flowChartAlternateProcess', 'flowChartDecision',
  'flowChartInputOutput', 'flowChartPredefinedProcess', 'flowChartInternalStorage',
  'flowChartDocument', 'flowChartTerminator', 'flowChartPreparation', 'flowChartManualInput',
  'flowChartConnector', 'flowChartSummingJunction',
  'flowChartMultidocument', 'flowChartManualOperation', 'flowChartOffpageConnector',
  'flowChartMagneticDisk', 'flowChartMagneticDrum', 'flowChartDisplay', 'flowChartDelay',
  'flowChartOr', 'flowChartCollate', 'flowChartSort', 'flowChartExtract', 'flowChartMerge',
  'flowChartPunchedTape',
] as const

const PRESETS = [...BASIC, ...RECTS, ...ARROWS, ...CALLOUTS, ...STARS_BANNERS, ...FLOWCHART] as const

/** The full flat vocabulary: 119 OOXML presets + the two structural specials. */
export type ShapeKind = (typeof PRESETS)[number] | 'line' | 'text'

export const SHAPE_KINDS: readonly ShapeKind[] = [...PRESETS, 'line', 'text']

/** Grouped for a real toolbar picker (<optgroup>) — a flat 121-item <select>
 *  is unusable. "Line" and "Text" get their own one-item group each, same
 *  as any other category, rather than a special case in the UI. */
export const SHAPE_CATEGORIES: readonly ShapeCategory[] = [
  { label: 'Basic', kinds: BASIC },
  { label: 'Rects', kinds: RECTS },
  { label: 'Arrows', kinds: ARROWS },
  { label: 'Callouts', kinds: CALLOUTS },
  { label: 'Stars & banners', kinds: STARS_BANNERS },
  { label: 'Flowchart', kinds: FLOWCHART },
  { label: 'Lines & text', kinds: ['line', 'text'] },
]

/** Human-readable label for a kind — camelCase prst names split on case
 *  boundaries and digit runs ("flowChartDecision" → "Flow Chart Decision",
 *  "star5" → "Star 5"). Good enough for a picker option and a fallback
 *  render's caption; not meant to match Office's exact display strings. */
export function shapeLabel(kind: string): string {
  const spaced = kind
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .replace(/([A-Za-z])(\d)/g, '$1 $2')
  return spaced.charAt(0).toUpperCase() + spaced.slice(1)
}

export interface ShapeSpec {
  id: string
  /** Stable identity returned by native XLSX hydration. It is intentionally
   * absent on browser-created shapes until the host saves and rehydrates. */
  readonly nativeIdentity?: NativeShapeIdentity
  kind: ShapeKind
  /** Text content — inline-editable for every kind except "line" and the
   *  arrow family (see hasText). */
  text?: string
  /** Fill color ('' or undefined = transparent). */
  fill?: string
  stroke?: string
  strokeWidth?: number
  textColor?: string
  fontSize?: number
}

/** One top-level shape in an XLSX worksheet drawing. Display names and anchor
 * indexes are mutable; the drawing part plus cNvPr id is the stable key. */
export interface NativeShapeIdentity {
  readonly drawingPart: string
  readonly objectId: number
}

const FILLED_DEFAULT = { fill: '#eef4f2', stroke: '#0fa98f', strokeWidth: 1.5, text: '' } as const
const ARROW_DEFAULT = { stroke: '#59636a', strokeWidth: 2 } as const
const CALLOUT_DEFAULT = { fill: '#fff8e6', stroke: '#f2a13c', strokeWidth: 1.5, text: 'Note' } as const

function buildDefaults(): Record<ShapeKind, Partial<ShapeSpec>> {
  const out = {} as Record<ShapeKind, Partial<ShapeSpec>>
  for (const k of BASIC) out[k] = FILLED_DEFAULT
  for (const k of RECTS) out[k] = FILLED_DEFAULT
  for (const k of ARROWS) out[k] = ARROW_DEFAULT
  for (const k of CALLOUTS) out[k] = CALLOUT_DEFAULT
  for (const k of STARS_BANNERS) out[k] = FILLED_DEFAULT
  for (const k of FLOWCHART) out[k] = FILLED_DEFAULT
  out.line = ARROW_DEFAULT
  out.text = { text: 'Text', textColor: '#1d2427', fontSize: 14 }
  return out
}

export const SHAPE_DEFAULTS: Record<ShapeKind, Partial<ShapeSpec>> = buildDefaults()

const ARROW_KIND_SET: ReadonlySet<string> = new Set(ARROWS)

/** True when a shape kind carries editable text. Every preset does except
 *  the arrow family (rendered as a solid glyph — see arrowFillKinds on the
 *  Go side — where an overlaid text layer reads poorly) and "line" (a bare
 *  connector). */
export function hasText(kind: ShapeKind): boolean {
  return kind !== 'line' && !ARROW_KIND_SET.has(kind)
}
