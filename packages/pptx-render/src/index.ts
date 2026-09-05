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
