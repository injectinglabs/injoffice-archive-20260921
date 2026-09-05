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

// ---- Text overflow defense (fixes a live-staging bug: a long title on the
// 'title' SlideKind rendered past the canvas's right edge) ----
//
// ROUND 4 FIX (WRONG/INCOMPLETE — kept here as history, corrected below):
// diagnosed as "Archivo" (boardroom's font) never being loaded as a web
// font, shipped canvasFontFamily's fallback + fitFontSizePx's shrink + a
// hard Group clip. Nick re-verified on staging and the bug was STILL
// present — AND reproduced identically on the "terra" theme, which uses
// the exact same displayFont as boardroom (every BUILTIN_THEME does — see
// themes.ts), so "which theme" was never actually the variable; the font-
// loading theory didn't explain why the fix didn't work, it just happened
// to point at a real (but insufficient) contributing factor. The live
// symptom, described precisely: text renders fully, bold, correctly
// sized — it just never wraps at all, running in ONE line all the way to
// the canvas's own edge with no visible clip boundary. That is NOT what a
// working clipped-Group + working word-wrap should ever produce (clipping
// would cut the line off well short of the canvas edge, at the shape's own
// ~81%-width box, with the background visibly continuing past it) — it is
// exactly what an UNWRAPPED, UNCLIPPED single line looks like. Conclusion:
// something about relying on Konva's OWN width-based wrap decision (which
// ultimately depends on the browser canvas's own text-measurement/font-
// resolution behavior) is not trustworthy here, in a way this package's
// headless node-canvas test environment does not reproduce and therefore
// cannot be fully diagnosed from here.
//
// ROUND 5 FIX: stop trusting Konva/canvas measurement for the WRAP
// DECISION at all. wrapLineToWidth (below) pre-computes line breaks itself,
// using the SAME coarse character-width heuristic qc.ts's
// estimateWrappedLines already uses for the DOM renderer — deterministic,
// testable, and independent of whatever font actually ends up resolving in
// whatever browser. DeckCanvasView now renders the ALREADY-WRAPPED text
// with wrap="none" (Konva still respects explicit '\n' breaks in any wrap
// mode; "none" only disables ITS OWN additional width-based breaking, which
// is exactly what's being replaced). The round-4 fixes are kept, now as
// secondary defense layers, not the primary mechanism:
//  1. canvasFontFamily's generic fallback — still a reasonable robustness
//     improvement for actual glyph rendering, just no longer load-bearing
//     for the wrap decision.
//  2. fitFontSizePx — still shrinks a single word too wide even for
//     wrapLineToWidth to break (word-wrap, manual or automatic, can't
//     split inside a word).
//  3. The hard Group clip (clipX/clipY/clipWidth/clipHeight) — still the
//     unconditional structural guarantee: nothing can ever PAINT outside a
//     shape's own box, regardless of ANY wrapping/measurement mechanism's
//     behavior in ANY browser.

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
