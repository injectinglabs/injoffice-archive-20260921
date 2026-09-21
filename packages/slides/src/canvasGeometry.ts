// Pure EMU→px geometry for DeckCanvasView — separated so the math is
// unit-testable independent of react-konva/canvas, same split
// @injoffice/shapes keeps between geometry.ts (tested) and ShapeFloat.tsx
// (rendering, untested directly).

import type { WireShapeKind } from './wire'

export const EMU_PER_INCH = 914400
export const EMU_PER_PT = 12700

export interface Viewport {
  scale: number
  widthPx: number
  heightPx: number
}

/** A uniform EMU→px scale that fits slideCx wide into fitWidthPx, height
 *  following the slide's own aspect ratio (never stretched). */
export function makeViewport(slideCx: number, slideCy: number, fitWidthPx: number): Viewport {
  const scale = slideCx > 0 ? fitWidthPx / slideCx : 0
  return { scale, widthPx: fitWidthPx, heightPx: slideCy * scale }
}

export function emuToPx(emu: number, scale: number): number {
  return emu * scale
}

/** Point size → px, through EMU (1pt = 12700 EMU) at the given EMU→px scale — keeps
 *  every dimension (position, size, stroke width, font size) on the same one scale factor. */
export function ptToPx(pt: number, scale: number): number {
  return pt * EMU_PER_PT * scale
}

export interface PxRect {
  x: number
  y: number
  w: number
  h: number
}

export function shapeRectPx(shape: { x: number; y: number; cx: number; cy: number }, scale: number): PxRect {
  return { x: emuToPx(shape.x, scale), y: emuToPx(shape.y, scale), w: emuToPx(shape.cx, scale), h: emuToPx(shape.cy, scale) }
}

/** Konva <RegularPolygon> side count for prst kinds that map onto a regular
 *  polygon (point-up, same convention @injoffice/shapes' regularPolygonPath
 *  uses) — null for kinds rendered some other way (ellipse/star/arrow/line/
 *  textBox each get their own Konva node type in DeckCanvasView). */
export function regularPolygonSides(kind: WireShapeKind): number | null {
  switch (kind) {
    case 'triangle':
      return 3
    case 'diamond':
      return 4
    case 'pentagon':
      return 5
    case 'hexagon':
      return 6
    default:
      return null
  }
}

/** Rounded-rect corner radius, same 0.16-of-shorter-side ratio
 *  @injoffice/shapes' roundRectPath uses, so a roundRect reads consistently
 *  whether it's drawn by the sheets shape layer or a slide. */
export function roundRectCornerRadiusPx(w: number, h: number): number {
  return Math.min(w, h) * 0.16
}

/** One flattened line of text for a textBox shape: bullet paragraphs get a
 *  glyph prefix, indented per level. DeckCanvasView renders a shape's whole
 *  paragraph list as ONE Konva Text node (see its doc comment for why) —
 *  this is the pure line-building half of that, independently testable. */
export function paragraphLine(runsText: string, bullet: boolean | undefined, level: number | undefined): string {
  if (!bullet) return runsText
  const lvl = level ?? 0
  const indent = '  '.repeat(lvl)
  const glyph = lvl % 2 === 0 ? '•' : '◦'
  return `${indent}${glyph} ${runsText}`
}

// ---- S8 (basic slice): drag/resize px↔EMU math ----
// The Konva-side glue (DeckCanvasView) wraps every shape in a <Group> whose
// x/y/width/height are TOP-LEFT px, uniformly, regardless of what anchor
// convention the inner primitive itself uses (Ellipse/Star/RegularPolygon
// are center-anchored in Konva; Rect/Line/Text are top-left) — so drag and
// resize always read/write the SAME simple rect, and this math never needs
// to special-case shape kind. These two functions are that rect's px→EMU
// half; the Group wrapping and the Konva event wiring are DeckCanvasView's
// job (untested directly, same split as the rest of this file).

/** Minimum on-canvas size (px) a drag/resize is allowed to shrink a shape
 *  to — floors both live interactive resize (Transformer's boundBoxFunc)
 *  and the committed result, so a shape can never be dragged to zero/
 *  negative size. */
export const MIN_SHAPE_PX = 8

/** A shape's new EMU x/y after a drag ends at newXPx/newYPx (top-left, in
 *  the SAME px space shapeRectPx produces at this scale). */
export function dragResultEmu(newXPx: number, newYPx: number, scale: number): { x: number; y: number } {
  if (scale <= 0) return { x: 0, y: 0 }
  return { x: Math.round(newXPx / scale), y: Math.round(newYPx / scale) }
}

/** A shape's new EMU size after a resize ends at newWidthPx/newHeightPx —
 *  floored at MIN_SHAPE_PX so a committed resize can't collapse a shape. */
export function resizeResultEmu(newWidthPx: number, newHeightPx: number, scale: number): { cx: number; cy: number } {
  if (scale <= 0) return { cx: 0, cy: 0 }
  return {
    cx: Math.round(Math.max(MIN_SHAPE_PX, newWidthPx) / scale),
    cy: Math.round(Math.max(MIN_SHAPE_PX, newHeightPx) / scale),
  }
}

// ---- Text overflow defense ----
// Precompute line breaks with the same coarse character-width heuristic as
// qc.ts. DeckCanvasView renders those explicit breaks with wrap="none", so
// the wrap decision is deterministic across browser font-loading states.
// This is approximate preview layout, not supplied-font text measurement.
// A generic font fallback, shrinking overlong words, and a shape-local Group
// clip provide additional rendering bounds. Headless Konva tests exercise
// integration but cannot reproduce every browser font-measurement failure.

const AVG_CHAR_WIDTH_RATIO = 0.55

/** Estimated on-screen width (px) of `text` set at fontSizePx — a coarse
 *  heuristic (real glyph widths vary by character/font), same honest scope
 *  as qc.ts's estimateWrappedLines; NOT a substitute for real canvas
 *  measureText, just a cheap, deterministic, testable safety check. */
export function estimateTextWidthPx(text: string, fontSizePx: number): number {
  return text.length * fontSizePx * AVG_CHAR_WIDTH_RATIO
}

/** Shrinks fontSizePx (never below minFontSizePx) until `text`'s LONGEST
 *  WORD is estimated to fit within boxWidthPx. Word-wrap only ever breaks
 *  at word boundaries — a long STRING just wraps to more lines (a
 *  fixed-width box handles that fine on its own), but a single word wider
 *  than the box can't wrap and WOULD visually overflow it. Returns
 *  fontSizePx unchanged for degenerate input (no text, no box width). */
export function fitFontSizePx(text: string, boxWidthPx: number, fontSizePx: number, minFontSizePx = 10): number {
  if (!text || boxWidthPx <= 0) return fontSizePx
  const longestWord = text.split(/\s+/).reduce((max, w) => (w.length > max.length ? w : max), '')
  let size = fontSizePx
  while (size > minFontSizePx && estimateTextWidthPx(longestWord, size) > boxWidthPx) {
    size -= 1
  }
  return size
}

/** A font-family value safe to hand a <canvas> context directly: `name`
 *  plus a guaranteed-available generic fallback, so canvas text resolution
 *  never depends on a specific (possibly unloaded) named font actually
 *  being available. "" / undefined → undefined (Konva's own default). */
export function canvasFontFamily(name: string | undefined): string | undefined {
  const trimmed = name?.trim()
  return trimmed ? `${trimmed}, sans-serif` : undefined
}

/** Greedily breaks ONE logical line of text into multiple lines that each
 *  fit boxWidthPx at fontSizePx, by our own estimateTextWidthPx heuristic —
 *  the primary fix for the round-5 corrected diagnosis (see this file's
 *  "Text overflow defense" comment above): DeckCanvasView renders the
 *  RESULT of this with Konva's wrap disabled, instead of asking Konva/the
 *  browser canvas to decide where to break. Any LEADING whitespace (a
 *  bullet's indent from paragraphLine, e.g. "  ◦ ") is preserved and
 *  re-applied to every wrapped continuation line, so an indented bullet
 *  still reads as one indented block after wrapping — internal spacing
 *  beyond the leading run is normalized to single spaces, not preserved
 *  exactly. A single word wider than boxWidthPx on its own still goes out
 *  as its own (overflowing) line — word wrap, manual or automatic, cannot
 *  break inside a word; fitFontSizePx is the defense for THAT case, and
 *  the hard Group clip is the unconditional backstop regardless. */
export function wrapLineToWidth(line: string, boxWidthPx: number, fontSizePx: number): string[] {
  if (!line.trim()) return [line]
  if (boxWidthPx <= 0 || fontSizePx <= 0) return [line]
  const indent = /^\s*/.exec(line)?.[0] ?? ''
  const words = line.slice(indent.length).split(/\s+/).filter(Boolean)
  const out: string[] = []
  let current = ''
  for (const word of words) {
    const candidate = current ? `${current} ${word}` : word
    if (current && estimateTextWidthPx(indent + candidate, fontSizePx) > boxWidthPx) {
      out.push(indent + current)
      current = word
    } else {
      current = candidate
    }
  }
  if (current) out.push(indent + current)
  return out.length ? out : [line]
}
