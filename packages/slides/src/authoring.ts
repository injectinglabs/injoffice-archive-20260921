// Pure authoring boundary. Keep this entry free of React, Konva, DOM, CSS,
// browser measurement, and legacy preview components. Native Office consumers
// import this subpath instead of evaluating @injoffice/slides' UI entry.
export { compileDeckToWire, compileSlide, splitInlineRuns } from './compile'
export { compileDiagramShapes } from './diagram'
export type { DiagramRegion } from './diagram'
export { BUILTIN_THEMES, boardroomTheme, midnightTheme, slateTheme, terraTheme, forestTheme, plumTheme, resolveTheme } from './themes'
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
