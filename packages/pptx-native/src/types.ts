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
export interface NativeTransform { x: number; y: number; cx: number; cy: number; quarterTurns?: 1 | 2 | 3 }

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
  horizontalOverflow: 'overflow'
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

export interface NativeOpaqueChart {
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
  /** Absent only when compatibility.status is refused and rendering must use a placeholder. */
  preset?: NativeShapePreset
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
