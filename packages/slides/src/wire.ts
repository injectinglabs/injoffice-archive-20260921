// Wire format for @injoffice/slides' compiled Deck — the JSON shape a
// DeckSpec compiles INTO (compile.ts) and what a host's gateway can accept
// in a build-from-request endpoint (a POST body) to produce a real .pptx
// via go/pptxpatch server-side. Field names/shapes mirror go/pptxpatch's
// Deck/Slide/Shape/Paragraph/TextRun 1:1 — the same "wire mirrors the Go
// write struct" convention @injoffice/charts' toFile.ts (WireChart) and
// @injoffice/shapes' toFile.ts (WireShape) already established for xlsx.
// Every field here is EMU-native (X/Y/Cx/Cy), matching pptxpatch directly —
// no px/pt conversion at this boundary.

export type WirePlaceholder = '' | 'title' | 'ctrTitle' | 'subTitle' | 'body'

/** Kind IS the OOXML prst value (pptxpatch's own convention) — keep in sync
 *  with go/pptxpatch's presetGeomKinds. */
export type WireShapeKind = 'textBox' | 'line' | 'rect' | 'roundRect' | 'ellipse' | 'triangle' | 'diamond' | 'rightArrow' | 'pentagon' | 'hexagon' | 'star5'

export type WireAlign = '' | 'l' | 'ctr' | 'r'

export interface WireTextRun {
  text: string
  bold?: boolean
  italic?: boolean
  sizePt?: number
  color?: string
  font?: string
}

export interface WireParagraph {
  runs: WireTextRun[]
  align?: WireAlign
  level?: number
  bullet?: boolean
}

export interface WireShape {
  kind: WireShapeKind
  placeholder?: WirePlaceholder
  name?: string
  x: number
  y: number
  cx: number
  cy: number
  fill?: string
  stroke?: string
  strokeWidthPt?: number
  paragraphs?: WireParagraph[]
  /** Only meaningful for kind 'line': draw a triangle arrowhead at the
   *  connector's first/second a:xfrm point respectively. Added for S12's
   *  diagram slice (process-flow arrows) — see go/pptxpatch's Shape.HeadArrow/
   *  TailArrow doc comment, which this mirrors 1:1. */
  headArrow?: boolean
  tailArrow?: boolean
  /** Only meaningful for kind 'line': mirrors a:xfrm's own flipH — a straight
   *  connector always draws its OWN bounding box's top-left-to-bottom-right
   *  diagonal unless flipped. See go/pptxpatch's Shape.FlipH doc comment
   *  (mirrored 1:1 here) for the derivation; @injoffice/slides' diagram.ts
   *  computes this for org-chart connectors. */
  flipH?: boolean
  /** compile.ts-only bookkeeping: this shape's STABLE ROLE within its slide
   *  kind (e.g. 'title', 'bullets'), the same key SlideSpec.shapeOverrides
   *  is keyed by (see types.ts) — how DeckCanvasView's drag/resize editor
   *  reports WHICH shape moved back to a host, without relying on array
   *  index (which shifts as content changes). Round-trips harmlessly to
   *  go/pptxpatch's BuildPPTX (an unrecognized JSON field is ignored) but
   *  carries no OOXML meaning there — purely a client-side identity. */
  key?: string
  /** This shape's entrance animation (S11) — mirrors go/pptxpatch's
   *  ShapeAnimation 1:1, flattened the same way fill/stroke are (not
   *  nested), matching this file's own convention. Omitted = no animation,
   *  same as every other optional field here. */
  enter?: 'fade' | 'flyIn'
  /** 'flyIn' only. */
  enterDirection?: 'left' | 'right' | 'top' | 'bottom'
  enterDelayMs?: number
  enterDurationMs?: number
  /** 'flyIn' only. */
  enterDistance?: number
}

export interface WireSlide {
  shapes: WireShape[]
  background?: string
  /** This slide's entrance transition (S11) — mirrors go/pptxpatch's
   *  SlideTransition 1:1. Omitted = no transition. */
  transition?: 'fade' | 'push' | 'wipe'
  /** 'push'/'wipe' only. */
  transitionDirection?: 'left' | 'right' | 'top' | 'bottom'
}

export interface WireDeck {
  slides: WireSlide[]
  /** Slide size in EMU; omitted = the gateway's 16:9 default (12192000×6858000). */
  cx?: number
  cy?: number
}
