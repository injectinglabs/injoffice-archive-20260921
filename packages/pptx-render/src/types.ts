import type {
  NativeArrowEnd,
  NativeCompatibilityStatus,
  NativeDiagnosticSeverity,
  NativePictureCrop,
  NativeShapePreset,
  NativeTextAlign,
} from '@injoffice/pptx-native'
import type {
  NativeFontManifest,
  NativeFontResolver,
  NativeTextDecision,
  NativeTextShaper,
  OpenTypeFeature,
  FontVariation,
  ShapedGlyph,
  ShapedCluster,
  TextDirection,
} from '@injoffice/font-metrics/layout'

export const PPTX_RENDER_TREE_VERSION = 'pptx-render-tree/v2' as const

export const PPTX_RENDER_LIMITS = Object.freeze({
  maxDepth: 64,
  maxNodes: 250_000,
  maxPathCommands: 256,
  maxPaintCommands: 1_000_000,
  maxGlyphs: 2_000_000,
  maxClusters: 2_000_000,
  maxProviderDecisions: 100_000,
  maxProviderAttemptedFaceIds: 100_000,
  maxProviderResolveCalls: 250_000,
  maxProviderLoadCalls: 256,
  maxProviderShapeCalls: 250_000,
  maxCachedFontResources: 256,
  maxTextLines: 100_000,
  maxTextFragments: 250_000,
  maxFontResourceBytes: 67_108_864,
  maxUniqueFontBytes: 134_217_728,
  maxCoordinateEmu: 281_474_976_710_655,
  maxAffinePpm: Number.MAX_SAFE_INTEGER,
  maxReferencedAssetBytes: 536_870_912,
})

export interface RenderRect { readonly x: number; readonly y: number; readonly cx: number; readonly cy: number }

/** Integer affine transform. Coefficients are parts-per-million; translation is EMU. */
export interface RenderTransform {
  readonly aPpm: number
  readonly bPpm: number
  readonly cPpm: number
  readonly dPpm: number
  readonly txEmu: number
  readonly tyEmu: number
}

export type RenderClip = { readonly kind: 'rect'; readonly rect: RenderRect } | { readonly kind: 'roundRect'; readonly rect: RenderRect; readonly radiusEmu: number }

export type RenderPathCommand =
  | { readonly kind: 'moveTo'; readonly x: number; readonly y: number }
  | { readonly kind: 'lineTo'; readonly x: number; readonly y: number }
  | { readonly kind: 'rect'; readonly rect: RenderRect }
  | { readonly kind: 'roundRect'; readonly rect: RenderRect; readonly radiusEmu: number }
  | { readonly kind: 'ellipse'; readonly rect: RenderRect }
  | { readonly kind: 'close' }

export interface RenderPaint { readonly color: string }
export interface RenderStroke extends RenderPaint {
  readonly widthEmu: number
  /** Required on native shape/connector strokes; legacy table borders omit these fields. */
  readonly cap?: 'flat' | 'round' | 'square'
  readonly join?: 'round' | 'bevel' | 'miter'
  readonly dash?: 'solid'
  readonly miterLimit?: number
}

export type RenderDiagnosticCode =
  | 'native.compatibility'
  | 'native.refused'
  | 'text.refused'
  | 'text.overflow'
  | 'text.layoutMetadataUnavailable'
  | 'text.bidiUnavailable'
  | 'text.verticalUnsupported'
  | 'text.verticalAnchorUnavailable'
  | 'text.deterministicLayout'
  | 'text.metricsUnavailable'
  | 'text.wrapUnavailable'
  | 'text.paragraphSemanticsUnavailable'
  | 'text.inheritanceUnavailable'
  | 'text.providerBudget'
  | 'chart.missingPreview'
  | 'asset.hostResolutionRequired'
  | 'render.preserveOnly'

export interface RenderDiagnostic {
  readonly severity: NativeDiagnosticSeverity
  readonly code: RenderDiagnosticCode | string
  readonly message: string
  readonly slideId: string
  readonly elementId?: string
  readonly sourceCode?: string
}

export interface RenderGlyph extends ShapedGlyph {
  readonly xEmu: number
  readonly yEmu: number
  readonly advanceXEmu: number
  readonly advanceYEmu: number
  readonly offsetXEmu: number
  readonly offsetYEmu: number
}

export interface RenderCluster extends Omit<ShapedCluster, 'startUtf16' | 'endUtf16'> {
  /** UTF-16 offsets into the owning RenderTextRunNode.text fragment. */
  readonly startUtf16: number
  readonly endUtf16: number
  readonly advanceInlineEmu: number
}

export interface RenderTextRunNode {
  /** Marker UTF-16 offsets refer to paragraph.bulletCharacter, not a content run. */
  readonly sourceRole?: 'paragraphBullet'
  readonly kind: 'textRun'
  readonly sourceElementId: string
  readonly paragraphIndex: number
  readonly runIndex: number
  /** Original source-run UTF-16 range represented by this (possibly wrapped) fragment. */
  readonly startUtf16: number
  readonly endUtf16: number
  readonly text: string
  readonly direction: TextDirection
  readonly fontSizeMilliPoints: number
  readonly faceId?: string
  readonly contentDigest?: string
  readonly color: string
  readonly bold: boolean
  readonly italic: boolean
  readonly x: number
  readonly baselineY: number
  readonly advanceInlineEmu: number
  readonly lineHeightEmu: number
  readonly glyphs: readonly RenderGlyph[]
  readonly clusters: readonly RenderCluster[]
  readonly decisions: readonly NativeTextDecision[]
  readonly attemptedFaceIds: readonly string[]
  readonly status: 'shaped' | 'refused'
}

export interface RenderParagraphNode {
  readonly marker?: RenderTextRunNode
  readonly kind: 'paragraph'
  readonly sourceElementId: string
  readonly paragraphIndex: number
  readonly lineIndex: number
  readonly align: NativeTextAlign
  readonly direction: TextDirection
  readonly level: number
  readonly bullet: boolean
  readonly x: number
  readonly y: number
  readonly widthEmu: number
  readonly heightEmu: number
  readonly runs: readonly RenderTextRunNode[]
  /** Authored U+0020 ranges consumed only because this soft wrap was taken. */
  readonly consumedSoftSeparators?: readonly {
    readonly sourceElementId: string
    readonly paragraphIndex: number
    readonly runIndex: number
    readonly startUtf16: number
    readonly endUtf16: number
  }[]
}

export interface RenderTextBodyNode {
  readonly kind: 'textBody'
  readonly sourceElementId: string
  readonly bounds: RenderRect
  /** Text-only physical mapping; parent shape/group transforms remain separate. */
  readonly transform?: RenderTransform
  readonly fidelity: 'native' | 'deterministicNative' | 'approximateSourceFrame' | 'nativeUnavailable' | 'legacyUnavailable'
  /** Explicit InjOffice line-box policy; does not attest Office visual parity. */
  readonly lineLayoutPolicy?: 'max-run-natural-v1'
  readonly wrap?: 'square' | 'none'
  readonly verticalAnchor?: 'top' | 'center' | 'bottom'
  readonly autoFit?: 'none' | 'shape-source-frame'
  readonly horizontalOverflow?: 'overflow'
  readonly verticalOverflow?: 'overflow'
  readonly status: 'laidOut' | 'refused'
  readonly paragraphs: readonly RenderParagraphNode[]
  readonly refusalLabel?: string
}

interface RenderNodeBase {
  readonly sourceElementId: string
  readonly sourceKind: string
  readonly zIndex: number
  readonly transform: RenderTransform
  readonly bounds: RenderRect
  readonly clip?: RenderClip
  readonly compatibility: NativeCompatibilityStatus
}

export interface RenderShapeNode extends RenderNodeBase {
  readonly kind: 'shape'
  readonly preset: NativeShapePreset
  readonly path: readonly RenderPathCommand[]
  readonly fill?: RenderPaint
  readonly stroke?: RenderStroke
  readonly textBody?: RenderTextBodyNode
}

export interface RenderTextNode extends RenderNodeBase {
  readonly kind: 'text'
  readonly textBody: RenderTextBodyNode
}

export interface RenderConnectorNode extends RenderNodeBase {
  readonly kind: 'connector'
  readonly path: readonly RenderPathCommand[]
  readonly stroke?: RenderStroke
  readonly headArrow: boolean
  readonly tailArrow: boolean
  readonly headEnd?: Readonly<NativeArrowEnd>
  readonly tailEnd?: Readonly<NativeArrowEnd>
}

export interface RenderImageNode extends RenderNodeBase {
  readonly kind: 'image'
  readonly role: 'picture' | 'chartPreview'
  readonly assetId: string
  /** Exact DrawingML source-edge insets; adapters must crop before scaling. */
  readonly crop?: Readonly<NativePictureCrop>
  readonly contentType: string
  readonly sha256: string
  readonly byteLength: number
  readonly resolutionSource: 'sourceDeck' | 'host'
}

interface RenderTableCellNodeBase {
  readonly kind: 'tableCell'
  readonly sourceElementId: string
  readonly rowIndex: number
  readonly columnIndex: number
  readonly bounds: RenderRect
  readonly fill?: RenderPaint
  readonly border?: RenderStroke
}

export type RenderTableCellNode = RenderTableCellNodeBase & (
  | { readonly paragraph: RenderParagraphNode; readonly textBody?: never }
  | { readonly paragraph?: never; readonly textBody: RenderTextBodyNode }
)

export interface RenderTableNode extends RenderNodeBase {
  readonly kind: 'table'
  readonly rows: number
  readonly columns: number
  readonly cells: readonly RenderTableCellNode[]
}

export interface RenderGroupNode extends RenderNodeBase {
  readonly kind: 'group'
  readonly children: readonly RenderNode[]
}

export interface RenderPlaceholderNode extends RenderNodeBase {
  readonly kind: 'placeholder'
  readonly reason: 'refused' | 'preserveOnly' | 'missingPreview' | 'textRefusal'
  readonly label: string
}

export type RenderNode =
  | RenderShapeNode
  | RenderTextNode
  | RenderConnectorNode
  | RenderImageNode
  | RenderTableNode
  | RenderGroupNode
  | RenderPlaceholderNode

export interface RenderAsset {
  readonly id: string
  readonly contentType: string
  readonly sha256: string
  readonly byteLength: number
  readonly resolutionSource: 'sourceDeck' | 'host'
}

export interface SlideRenderTree {
  readonly version: typeof PPTX_RENDER_TREE_VERSION
  readonly documentId: string
  readonly slideId: string
  readonly slideIndex: number
  readonly size: { readonly cx: number; readonly cy: number }
  readonly background: RenderPaint
  readonly clip: RenderClip
  readonly nodes: readonly RenderNode[]
  readonly assets: readonly RenderAsset[]
  readonly diagnostics: readonly RenderDiagnostic[]
}

export interface NativePptxTextDefaults {
  readonly fontFamilies: readonly string[]
  readonly fontSizeHundredthPt: number
  readonly script: string
  readonly language: string
  readonly direction: TextDirection
  readonly fallbackChainIds?: readonly string[]
}

export interface NativePptxTextRunContext {
  readonly slideId: string
  readonly elementId: string
  readonly elementKind: 'text' | 'shape' | 'table'
  readonly paragraphIndex: number
  readonly runIndex: number
  readonly text: string
}

export interface NativePptxTextOverride {
  readonly script?: string
  readonly language?: string
  readonly direction?: TextDirection
  readonly fontFamilies?: readonly string[]
  readonly fallbackChainIds?: readonly string[]
  readonly features?: readonly OpenTypeFeature[]
  readonly variations?: readonly FontVariation[]
  readonly letterSpacingMilliPoints?: number
  readonly wordSpacingMilliPoints?: number
}

export interface NativePptxTextLayout {
  readonly manifest: NativeFontManifest
  readonly resolver: NativeFontResolver
  readonly shaper: NativeTextShaper
  readonly defaults: NativePptxTextDefaults
  /** Host-supplied language/script/direction metadata. The core never infers these from characters. */
  readonly resolveRun?: (context: NativePptxTextRunContext) => NativePptxTextOverride
}

export interface CompileSlideOptions {
  readonly textLayout: NativePptxTextLayout
  /** Read-only saved-frame preview of explicitly marked spAutoFit projections; never resizes or qualifies Office fidelity. */
  readonly sourceFrameAutoFitPreview?: boolean
  /** Opt into measured mixed-run line boxes and anchors, labeled deterministicNative. Omission retains strict qualification. */
  readonly lineLayoutPolicy?: 'max-run-natural-v1'
  readonly maxDepth?: number
  readonly maxNodes?: number
  readonly maxGlyphs?: number
  readonly maxClusters?: number
  readonly maxCoordinateEmu?: number
}

export class RenderCompileError extends Error {
  readonly code: string
  readonly path: string

  constructor(code: string, path: string, message: string) {
    super(`${path}: ${message}`)
    this.name = 'RenderCompileError'
    this.code = code
    this.path = path
  }
}
