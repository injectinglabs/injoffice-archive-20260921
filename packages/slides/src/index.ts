export { DeckView, SlideView, renderInline } from './DeckView'
export { deckFromOutline, resetSlideIds } from './outline'
export { DeckEditorPanel } from './DeckEditorPanel'
export { addSlide, removeSlide, moveSlide, updateSlide, bulletsToText, textToBullets } from './edit'
export { BUILTIN_THEMES, boardroomTheme, midnightTheme, slateTheme, terraTheme, forestTheme, plumTheme, resolveTheme } from './themes'
export { auditDeck, auditSlide, estimateWrappedLines, formatDeckAudit } from './qc'
export type { QCIssue, QCIssueKind } from './qc'
export { compileDeckToWire, compileSlide, splitInlineRuns } from './compile'
export { compileDiagramShapes } from './diagram'
export type { DiagramRegion } from './diagram'
export { DeckCanvasView, useCompiledDeck } from './DeckCanvasView'
export type { DeckCanvasViewProps, ShapeEditPatch } from './DeckCanvasView'
export type {
  WireAlign,
  WireDeck,
  WireParagraph,
  WirePlaceholder,
  WireShape,
  WireShapeKind,
  WireSlide,
  WireTextRun,
} from './wire'
export type {
  AnimDirection,
  DeckSpec,
  DeckTheme,
  DiagramSpec,
  EnterAnimation,
  OrgChartNode,
  ShapeAnimationSpec,
  ShapePositionOverride,
  SlideKind,
  SlideSpec,
  SlideTransitionKind,
  SlideTransitionSpec,
} from './types'
