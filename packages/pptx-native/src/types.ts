import type {NativeLiteralArea} from './chartAreaTypes.js'
import type {
  NativeAnimationEffect,
  NativeCompatibilityStatus,
  NativeDiagnosticSeverity,
  NativeDirection,
  NativeElementKind,
  NativeOrigin,
  NativeProvenance,
  NativePassthroughDisposition,
  NativePlaceholderType,
  NativeShapePreset,
  NativeStrokeCap,
  NativeStrokeDash,
  NativeStrokeJoin,
  NativeTextAlign,
  NativeTextVerticalAnchor,
  NativeTextWrap,
  NativeTransitionType,
} from './schema.generated'

export interface NativePptxDeck {
  contractVersion: 'pptx-native/v1'
  documentId: string
  origin: NativeOrigin
  sourceRevision?: string
  size: NativeSize
  assets: NativeAsset[]
  slides: NativeSlide[]
  compatibility: NativeCompatibility
}

export interface NativeSize { cx: number; cy: number }
export interface NativeTransform { x: number; y: number; cx: number; cy: number; quarterTurns?: 1 | 2 | 3; rotationAngle?: number; flipH?: boolean; flipV?: boolean }

export interface NativeSourceAnchor {
  partName: string
  objectId: string
  relationshipId?: string
  fingerprintSha256: string
}

export interface NativePassthroughRef {
  token: string
  ownerPart: string
  fingerprintSha256: string
  disposition: NativePassthroughDisposition
}

export interface NativeDiagnosticScope {
  slideId?: string
  elementId?: string
  partName?: string
}

export interface NativeDiagnostic {
  severity: NativeDiagnosticSeverity
  code: string
  message: string
  scope?: NativeDiagnosticScope
}

export interface NativeCompatibility {
  status: NativeCompatibilityStatus
  diagnostics: NativeDiagnostic[]
}

export interface NativeAsset {
  id: string
  provenance: NativeProvenance
  contentType: string
  sha256: string
  byteLength: number
  dataBase64?: string
  source?: NativeSourceAnchor
  passthrough: NativePassthroughRef[]
}

export interface NativeTextRun {
  text: string
  bold?: boolean
  italic?: boolean
  fontSizeHundredthPt?: number
  kerningThresholdHundredthPt?: number
  color?: string
  fontFamily?: string
  language?: string
}

export interface NativeParagraph {
  runs: NativeTextRun[]
  align?: NativeTextAlign
  level?: number
  bullet?: boolean
  /** One authored Unicode marker. Native exact marker/font layout may still refuse. */
  bulletCharacter?: string
  /** Exact authored marker family; never a Unicode replacement or fallback. */
  bulletFontFamily?: string
  /** Source charset=2; requires independently qualified font cmap evidence. */
  bulletFontEncoding?: 'windows-symbol-byte-v1'
  marginLeftEmu?: number
  indentEmu?: number
}

/** Exact v1 horizontal text-frame slice with materialized OOXML defaults. */
export interface NativeTextBodyLayout {
  leftInsetEmu: number
  rightInsetEmu: number
  topInsetEmu: number
  bottomInsetEmu: number
  wrap: NativeTextWrap
  verticalAnchor: NativeTextVerticalAnchor
  autoFit: 'none' | 'shape-source-frame'
  horizontalOverflow: 'overflow' | 'clip'
  verticalOverflow: 'overflow'
  writingMode?: 'vertical-clockwise'
}

export interface NativeStroke {
  color: string
  widthEmu: number
  cap?: NativeStrokeCap
  join?: NativeStrokeJoin
  dash?: NativeStrokeDash
  miterLimit?: number
}

export interface NativeAnimation {
  effect: NativeAnimationEffect
  direction?: NativeDirection
  delayMs?: number
  durationMs?: number
  distancePpm?: number
}

export interface NativeTransition { type: NativeTransitionType; direction?: NativeDirection }

export interface NativeTableBorder { color: string; widthEmu: number }

export interface NativeTableCell {
  text: string
  /** Present with textBody only for self-contained native DrawingML cell text. */
  paragraphs?: NativeParagraph[]
  /** Present with paragraphs only; legacy authored cells retain text/align. */
  textBody?: NativeTextBodyLayout
  fill?: string
  border?: NativeTableBorder
  align?: NativeTextAlign
}

export interface NativeTable {
  columnWidths: number[]
  rowHeights: number[]
  rows: NativeTableCell[][]
}

export interface NativeLiteralPie {
  profile: 'literal-pie-v1'
  firstSliceAngle: number
  values: number[]
  colors: string[]
}

export interface NativeLiteralDoughnut {
  profile: 'literal-doughnut-v1'
  firstSliceAngle: number
  holeSize: number
  values: number[]
  colors: string[]
}

export interface NativeChartAxisLabelStyle {
 fontFamily:string
 fontSize:number
 color:string
 bold:boolean
 italic:boolean
 language:string
}
export interface NativeChartAxisLabels {
 profile:'explicit-axis-labels-v1'
 position:'low'|'high'
 majorTickMark:'none'|'out'
 style:NativeChartAxisLabelStyle
 majorUnit?:string
 numberFormat?:string
}

export interface NativeLiteralBarAxis {
 labels?:NativeChartAxisLabels
 id:number
 crossAxisId:number
 orientation:'minMax'|'maxMin'
 position:'b'|'l'
 deleted:boolean
 color?:string
 widthEmu?:number
 min?:string
 max?:string
 crossesAt?:string
}
export interface NativeLiteralBarSeries {
 index:number
 order:number
 title?:string
 values:string[]
 colors:string[]
}
export interface NativeLiteralBar {
 profile:'literal-bar-v1'
 barDirection:'column'|'bar'
 grouping:'clustered'
 dataOrigin:'literal'
 gapWidth:number
 overlap:0
 categories:string[]
 series:NativeLiteralBarSeries[]
 categoryAxis:NativeLiteralBarAxis
 valueAxis:NativeLiteralBarAxis
}
export interface NativeLiteralConnectedSeries {
 index:number
 order:number
 title?:string
 values:string[]
 xValues?:string[]
 color:string
 widthEmu:number
}
/** Explicit straight source lines; categories are empty only for XY scatter. */
export interface NativeLiteralConnected {
 profile:'literal-line-v1'|'literal-scatter-v1'
 dataOrigin:'literal'
 categories:string[]
 series:NativeLiteralConnectedSeries[]
 xAxis:NativeLiteralBarAxis
 yAxis:NativeLiteralBarAxis
}
export interface NativeOpaqueChart {
 literalArea?:NativeLiteralArea
 literalConnected?:NativeLiteralConnected
 literalBar?:NativeLiteralBar
  literalDoughnut?: NativeLiteralDoughnut
  literalPie?: NativeLiteralPie
  chartPart: string
  relationshipId: string
  opaqueRef: NativePassthroughRef
  previewAssetId?: string
}

interface NativeElementBase {
  kind: NativeElementKind
  id: string
  provenance: NativeProvenance
  name?: string
  transform: NativeTransform
  animation?: NativeAnimation
  source?: NativeSourceAnchor
  passthrough: NativePassthroughRef[]
  compatibility: NativeCompatibility
}

export interface NativeTextElement extends NativeElementBase {
  kind: 'text'
  placeholder?: NativePlaceholderType
  paragraphs: NativeParagraph[]
  textBody?: NativeTextBodyLayout
}

export interface NativeShapeElement extends NativeElementBase {
  kind: 'shape'
  /** Exactly one of preset or geometry unless explicitly refused. */
  preset?: NativeShapePreset
  geometry?: NativeEvaluatedGeometry
  placeholder?: NativePlaceholderType
  fill?: string
  stroke?: NativeStroke
  paragraphs: NativeParagraph[]
  textBody?: NativeTextBodyLayout
}

/** Source DrawingML names. Omitted width/length stay omitted; not pixel geometry. */
export interface NativeArrowEnd {
  type: 'none' | 'triangle' | 'arrow' | 'stealth' | 'diamond' | 'oval'
  w?: 'sm' | 'med' | 'lg'
  len?: 'sm' | 'med' | 'lg'
}
export interface NativeConnectorElement extends NativeElementBase {
  kind: 'connector'
  stroke?: NativeStroke
  headArrow?: boolean
  tailArrow?: boolean
  headEnd?: NativeArrowEnd
  tailEnd?: NativeArrowEnd
  flipH?: boolean
}

/** DrawingML source-edge insets in 1/1000 percent (100000 = full image).
 * Opposing inset sums must be less than 100000. No image bytes are rewritten. */
export interface NativePictureCrop {
  left: number
  top: number
  right: number
  bottom: number
}

export interface NativePictureElement extends NativeElementBase {
  kind: 'picture'
  assetId: string
  crop?: NativePictureCrop
  /** Exact DrawingML roundRect preset with its default (empty avLst) adjustment. */
  clip?: 'roundRect'
}

export interface NativeTableElement extends NativeElementBase {
  kind: 'table'
  table: NativeTable
}

export interface NativeChartElement extends NativeElementBase {
  kind: 'chart'
  chart: NativeOpaqueChart
  source: NativeSourceAnchor
}

export interface NativeGroupElement extends NativeElementBase {
  kind: 'group'
  /** Original DrawingML chOff/chExt child coordinate space. Required for
   *  parsed groups; omission is the legacy authored identity-local form. */
  childTransform?: NativeTransform
  children: NativeElement[]
}

export type NativeElement =
  | NativeTextElement
  | NativeShapeElement
  | NativeConnectorElement
  | NativePictureElement
  | NativeTableElement
  | NativeChartElement
  | NativeGroupElement

export interface NativeSlide {
  id: string
  provenance: NativeProvenance
  background?: string
  transition?: NativeTransition
  elements: NativeElement[]
  source?: NativeSourceAnchor
  passthrough: NativePassthroughRef[]
  compatibility: NativeCompatibility
}

export type {
  NativeAnimationEffect,
  NativeCompatibilityStatus,
  NativeDiagnosticSeverity,
  NativeDirection,
  NativeElementKind,
  NativeOrigin,
  NativeProvenance,
  NativePassthroughDisposition,
  NativePlaceholderType,
  NativeShapePreset,
  NativeTextAlign,
  NativeTextVerticalAnchor,
  NativeTextWrap,
  NativeTransitionType,
} from './schema.generated'

/** Numeric DrawingML geometry. Preserve-only until a serializer is qualified. */
export interface NativeEvaluatedGeometry {
 profile: 'drawingml-paths-v1'
 textRect: NativeGeometryTextRect
 paths: NativeGeometryPath[]
}
export interface NativeGeometryTextRect { x:number; y:number; cx:number; cy:number }
export interface NativeGeometryPath { fillMode:'norm'|'none'|'darken'|'darkenLess'|'lighten'|'lightenLess'; stroke:boolean; commands:NativeGeometryCommand[] }
export interface NativeGeometryCommand {
 kind:'moveTo'|'lineTo'|'quadBezierTo'|'cubicBezierTo'|'arcTo'|'close'
 x?:number; y?:number; x1?:number; y1?:number; x2?:number; y2?:number
 rx?:number; ry?:number; largeArc?:boolean; clockwise?:boolean
}
