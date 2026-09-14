export {createNativeLiteralBubblePaths,BUBBLE_PREVIEW_POLICY,BUBBLE_PREVIEW_DISCLOSURE} from './literalBubble.js'
export type {LiteralBubbleVector} from './literalBubble.js'
export {createNativeLiteralAreaPaths} from './literalArea.js'
export type {LiteralAreaVector} from './literalArea.js'
export {createNativeLiteralLinePaths} from './literalLine.js'
export {createNativeLiteralScatterPaths} from './literalScatter.js'
export type {LiteralConnectedVector} from './literalLine.js'
export {createNativeLiteralBarPaths} from './literalBar.js'
export type {LiteralBarVector} from './literalBar.js'
export { createNativeLiteralDoughnutPaths } from './literalDoughnut.js'
export { createNativeLiteralPiePaths } from './literalPie.js'
export { compileNativePptxSlide, stringifySlideRenderTree } from './compile.js'
export { connectorPath, presetPath } from './geometry.js'
export {
  createRecordingPaintSurface,
  paintSlideRenderTree,
  paintSlideRenderTreeToCanvas2D,
  replayPaintCommandsToCanvas2D,
} from './paint.js'
export type { Canvas2DCommandAdapter, PaintCommand, PaintSurface, RecordingPaintSurface } from './paint.js'
export {
  PPTX_RENDER_LIMITS,
  PPTX_RENDER_TREE_VERSION,
  RenderCompileError,
} from './types.js'
export type {
  CompileSlideOptions,
  NativePptxTextDefaults,
  NativePptxGlyphExtents,
  NativePptxGlyphExtentsRequest,
  NativePptxTextLayout,
  NativePptxTextOverride,
  NativePptxTextRunContext,
  RenderAsset,
  RenderClip,
  RenderCluster,
  RenderConnectorNode,
  RenderDiagnostic,
  RenderDiagnosticCode,
  RenderGlyph,
  RenderGroupNode,
  RenderImageNode,
  RenderNode,
  RenderPaint,
  RenderParagraphNode,
  RenderPathCommand,
  RenderPlaceholderNode,
  RenderRect,
  RenderShapeNode,
  RenderStroke,
  RenderTableCellNode,
  RenderTableNode,
  RenderTextNode,
  RenderTextBodyNode,
  RenderTextRunNode,
  RenderTransform,
  SlideRenderTree,
} from './types.js'

export {renderTransformMatrix} from './sourceRenderTransform.js'
export {SourceAffineBudget} from './sourceAffine.js'
export type {SourceAffineTransport} from './sourceAffine.js'

export {createNativeWorkbookChartPaths} from './workbookChartPaths.js'

export {createNativeLiteralStackedBarPaths} from './literalStackedBar.js'
export type {LiteralStackedBarVector} from './literalStackedBar.js'
export {createNativeLiteralStackedLinePaths} from './literalStackedLine.js'
export type {LiteralStackedLineVector} from './literalStackedLine.js'
