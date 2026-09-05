import { compileDiagramShapes } from './diagram'
import { resolveTheme } from './themes'
import type { DeckSpec, DeckTheme, ShapeAnimationSpec, ShapePositionOverride, SlideSpec } from './types'
import type { WireAlign, WireDeck, WireParagraph, WireShape, WireSlide, WireTextRun } from './wire'

// DeckSpec → Deck compiler (S7 pivot wire-up). DeckSpec stays the authoring/
// AI-generation contract; this is the pure function that turns it into the
// wire format go/pptxpatch's Deck model consumes, the same "toWireX" role
// @injoffice/charts' toWireCharts and @injoffice/shapes' toWireShapes play
// for xlsx (see those files before touching this one). Unlike those, there
// is no "existing file being patched" here — DeckSpec has never had a real
// pptx representation before, so every field always compiles to something;
// nothing is skipped (unlike a chart type with no OOXML mapping), so this
// returns a plain WireDeck, not a {result, skipped} pair.
//
// Layout: these authored regions originated with DeckView.tsx's legacy 16:9
// layout. This compiler now supplies deterministic integer-EMU authoring input
// to @injoffice/pptx-authored; the HTML view is not an Office authority. It
// targets the same visual regions, translated to EMU inches on a fixed
// 13.333 x 7.5 in canvas (pptxpatch's own 16:9 default) —
// close to, but not pixel-identical with, DeckView's fluid cqw layout
// (details in each per-kind function below). It also mirrors the shape of a
// host's existing sandbox-based python-pptx deck exporter (the kind of
// script this pivot is meant to eventually let a host retire) closely on
// purpose: same regions, same per-kind structure — but with REAL bulleted
// paragraphs (a:buChar) instead of that style of script's plain-text
// "bullet" workaround, and real prstGeom autoshapes instead of
// MSO_SHAPE.RECTANGLE stand-ins for the accent bars.
//
// S8 persistence: every shape a compileXxx function produces gets a STABLE
// KEY (its role within the slide kind — 'title', 'bullets', ...), and
// compileSlide applies SlideSpec.shapeOverrides on top of the computed
// default for any key present there — the same layout-editing round trip
// DeckCanvasView's drag/resize commits INTO (see types.ts's
// ShapePositionOverride and DeckCanvasView's onShapeOverride). Keys are
// per-KIND, not globally unique — a 'title' key means something different
// on a 'title' slide vs a 'section' slide, which is fine since overrides
// are looked up per-slide, never across slides.

const EMU_PER_INCH = 914400
function inch(n: number): number {
  return Math.round(n * EMU_PER_INCH)
}

/** First font-family in a CSS stack, quotes stripped (theme fonts are CSS
 *  stacks like "'Archivo', 'Helvetica Neue', Arial, sans-serif"). */
function firstFamily(stack: string, fallback: string): string {
  const fam = stack.split(',')[0]?.trim().replace(/^['"]|['"]$/g, '')
  return fam || fallback
}

interface RunStyle {
  sizePt?: number
  color?: string
  font?: string
  bold?: boolean
  italic?: boolean
}

const EMPH_RE = /\*\*([^*]+)\*\*|\*([^*]+)\*/g

/** Same bold/italic emphasis markup DeckView's renderInline consumes,
 *  split into real styled runs instead of React nodes. */
export function splitInlineRuns(text: string, base: RunStyle): WireTextRun[] {
  if (!text) return []
  const runs: WireTextRun[] = []
  let last = 0
  let m: RegExpExecArray | null
  EMPH_RE.lastIndex = 0
  const push = (t: string, extraBold: boolean, extraItalic: boolean) => {
    if (!t) return
    const run: WireTextRun = { text: t, ...base }
    if (base.bold || extraBold) run.bold = true
    if (base.italic || extraItalic) run.italic = true
    runs.push(run)
  }
  while ((m = EMPH_RE.exec(text)) !== null) {
    if (m.index > last) push(text.slice(last, m.index), false, false)
    if (m[1] !== undefined) push(m[1], true, false)
    else push(m[2] ?? '', false, true)
    last = m.index + m[0].length
  }
  if (last < text.length) push(text.slice(last), false, false)
  return runs.length ? runs : [{ text: '', ...base }]
}

function para(text: string, style: RunStyle, opts: { align?: WireAlign } = {}): WireParagraph {
  const p: WireParagraph = { runs: splitInlineRuns(text, style) }
  if (opts.align) p.align = opts.align
  return p
}

function bulletPara(text: string, style: RunStyle, level = 0): WireParagraph {
  const p: WireParagraph = { runs: splitInlineRuns(text, style), bullet: true }
  if (level) p.level = level
  return p
}

function textBox(x: number, y: number, w: number, h: number, paragraphs: WireParagraph[], opts: Partial<WireShape> = {}): WireShape {
  return { kind: 'textBox', x: inch(x), y: inch(y), cx: inch(w), cy: inch(h), paragraphs, ...opts }
}

function bar(x: number, y: number, w: number, h: number, fill: string, key: string): WireShape {
  return { kind: 'rect', x: inch(x), y: inch(y), cx: inch(w), cy: inch(h), fill, key }
}

function fonts(theme: DeckTheme) {
  return {
    display: firstFamily(theme.displayFont, 'Calibri Light'),
    body: firstFamily(theme.bodyFont, 'Calibri'),
    mono: firstFamily(theme.monoFont, 'Consolas'),
  }
}

function eyebrowShape(text: string | undefined, theme: DeckTheme, y: number): WireShape | null {
  if (!text) return null
  const f = fonts(theme)
  return textBox(0.9, y, 11.5, 0.4, [para(text.toUpperCase(), { sizePt: 12, color: theme.accent, font: f.mono })], { key: 'eyebrow' })
}

function pageNumberShape(index: number, total: number, theme: DeckTheme): WireShape {
  const f = fonts(theme)
  return textBox(12.1, 6.95, 1.0, 0.4, [para(`${index + 1} / ${total}`, { sizePt: 10, color: theme.muted, font: f.mono }, { align: 'r' })], { key: 'page-number' })
}

function bulletsBlock(items: string[] | undefined, x: number, y: number, w: number, h: number, theme: DeckTheme, key: string, colTitle?: string): WireShape {
  const f = fonts(theme)
  const paragraphs: WireParagraph[] = []
  if (colTitle) paragraphs.push(para(colTitle.toUpperCase(), { sizePt: 12, color: theme.accent, font: f.mono }))
  for (const item of items ?? []) paragraphs.push(bulletPara(item, { sizePt: 16, color: theme.ink, font: f.body }))
  return textBox(x, y, w, h, paragraphs, { key })
}

function compileTitle(s: SlideSpec, theme: DeckTheme): WireShape[] {
  const f = fonts(theme)
  const shapes: WireShape[] = []
  const eyebrow = eyebrowShape(s.eyebrow, theme, 2.0)
  if (eyebrow) shapes.push(eyebrow)
  shapes.push(
    textBox(0.9, 2.5, 10.8, 2.0, [para(s.title ?? '', { sizePt: 40, color: theme.ink, font: f.display, bold: true })], {
      placeholder: 'title',
      name: 'Title',
      key: 'title',
    }),
  )
  if (s.subtitle) shapes.push(textBox(0.9, 4.4, 8.5, 1.4, [para(s.subtitle, { sizePt: 20, color: theme.muted, font: f.body })], { key: 'subtitle' }))
  if (s.body) shapes.push(textBox(0.9, 5.5, 8.5, 1.2, [para(s.body, { sizePt: 14, color: theme.muted, font: f.body })], { key: 'body' }))
  shapes.push(bar(0.9, 6.7, 1.5, 0.09, theme.accent, 'accent'))
  return shapes
}

function compileSection(s: SlideSpec, theme: DeckTheme): WireShape[] {
  const f = fonts(theme)
  const shapes: WireShape[] = []
  const eyebrow = eyebrowShape(s.eyebrow, theme, 2.2)
  if (eyebrow) shapes.push(eyebrow)
  shapes.push(
    textBox(0.9, 2.7, 10.8, 1.6, [para(s.title ?? '', { sizePt: 38, color: theme.ink, font: f.display, bold: true })], {
      placeholder: 'ctrTitle',
      name: 'Title',
      key: 'title',
    }),
  )
  const sub = s.subtitle ?? s.body
  if (sub) shapes.push(textBox(0.9, 4.2, 8.5, 1.6, [para(sub, { sizePt: 17, color: theme.muted, font: f.body })], { key: 'subtitle' }))
  return shapes
}

function compileQuote(s: SlideSpec, theme: DeckTheme): WireShape[] {
  const f = fonts(theme)
  const shapes: WireShape[] = [bar(1.2, 2.1, 10.9, 3.3, theme.surface, 'quote-panel'), bar(1.2, 2.1, 0.09, 3.3, theme.accent, 'quote-accent')]
  shapes.push(textBox(1.8, 2.5, 9.8, 2.0, [para(`“${s.quote ?? ''}”`, { sizePt: 24, color: theme.ink, font: f.body, italic: true })], { key: 'quote' }))
  if (s.subtitle) {
    shapes.push(textBox(1.8, 4.5, 9.0, 0.6, [para(`— ${s.subtitle.toUpperCase()}`, { sizePt: 12, color: theme.muted, font: f.mono })], { key: 'subtitle' }))
  }
  return shapes
}

function compileTwoCol(s: SlideSpec, theme: DeckTheme): WireShape[] {
  const f = fonts(theme)
  const shapes: WireShape[] = []
  const eyebrow = eyebrowShape(s.eyebrow, theme, 0.7)
  if (eyebrow) shapes.push(eyebrow)
  shapes.push(
    textBox(0.9, 1.1, 11.5, 1.0, [para(s.title ?? '', { sizePt: 28, color: theme.ink, font: f.display, bold: true })], { placeholder: 'title', name: 'Title', key: 'title' }),
  )
  shapes.push(bulletsBlock(s.bullets, 0.9, 2.3, 5.5, 4.7, theme, 'col-left', s.colTitles?.[0]))
  shapes.push(bulletsBlock(s.bulletsRight, 6.9, 2.3, 5.5, 4.7, theme, 'col-right', s.colTitles?.[1]))
  return shapes
}

function compileClosing(s: SlideSpec, theme: DeckTheme): WireShape[] {
  const f = fonts(theme)
  const shapes: WireShape[] = [
    textBox(1.5, 2.8, 10.3, 1.5, [para(s.title ?? 'Thank you', { sizePt: 40, color: theme.ink, font: f.display, bold: true }, { align: 'ctr' })], {
      placeholder: 'ctrTitle',
      name: 'Title',
      key: 'title',
    }),
  ]
  if (s.subtitle) shapes.push(textBox(1.5, 4.2, 10.3, 0.9, [para(s.subtitle, { sizePt: 18, color: theme.muted, font: f.body }, { align: 'ctr' })], { key: 'subtitle' }))
  shapes.push(bar(5.9, 5.3, 1.5, 0.09, theme.accent, 'accent'))
  return shapes
}

function compileBullets(s: SlideSpec, theme: DeckTheme): WireShape[] {
  const f = fonts(theme)
  const shapes: WireShape[] = []
  const eyebrow = eyebrowShape(s.eyebrow, theme, 0.7)
  if (eyebrow) shapes.push(eyebrow)
  if (s.title) shapes.push(textBox(0.9, 1.1, 11.5, 1.0, [para(s.title, { sizePt: 28, color: theme.ink, font: f.display, bold: true })], { placeholder: 'title', name: 'Title', key: 'title' }))
  let y = 2.2
  if (s.body) {
    shapes.push(textBox(0.9, y, 11.0, 1.0, [para(s.body, { sizePt: 15, color: theme.muted, font: f.body })], { key: 'body' }))
    y += 1.0
  }
  shapes.push({ ...bulletsBlock(s.bullets, 0.9, y, 11.4, 7.0 - y, theme, 'bullets'), placeholder: 'body', name: 'Body' })
  return shapes
}

/** Merge a SlideSpec's stored position/size override for `shape.key` (if
 *  any) on top of the compiler's own computed default — the only field
 *  compile.ts uses SlideSpec.shapeOverrides for. A shape without a key, or
 *  a key with no matching override, passes through unchanged. */
function applyOverride(shape: WireShape, overrides: Record<string, ShapePositionOverride> | undefined): WireShape {
  const o = shape.key ? overrides?.[shape.key] : undefined
  if (!o) return shape
  return {
    ...shape,
    x: o.x ?? shape.x,
    y: o.y ?? shape.y,
    cx: o.cx ?? shape.cx,
    cy: o.cy ?? shape.cy,
  }
}

/** Merge a SlideSpec's stored entrance-animation spec for `shape.key` (if
 *  any) onto the shape — the same "keyed by stable role, matched shapes
 *  only" pattern applyOverride uses for shapeOverrides (see types.ts's
 *  SlideSpec.shapeAnimations). A shape without a key, or a key with no
 *  matching animation, passes through unchanged (no `enter*` fields at
 *  all — the wire format's own "omitted = no animation" default). */
function applyAnimation(shape: WireShape, animations: Record<string, ShapeAnimationSpec> | undefined): WireShape {
  const spec = shape.key ? animations?.[shape.key] : undefined
  if (!spec) return shape
  return {
    ...shape,
    enter: spec.enter,
    enterDirection: spec.direction,
    enterDelayMs: spec.delayMs,
    enterDurationMs: spec.durationMs,
    enterDistance: spec.distance,
  }
}

/** Compile one SlideSpec into a WireSlide. Exported for direct testing and
 *  for hosts that want a single slide compiled without a whole DeckSpec. */
export function compileSlide(s: SlideSpec, theme: DeckTheme, index: number, total: number): WireSlide {
  let shapes: WireShape[]
  switch (s.kind) {
    case 'title':
      shapes = compileTitle(s, theme)
      break
    case 'section':
      shapes = compileSection(s, theme)
      break
    case 'quote':
      shapes = compileQuote(s, theme)
      break
    case 'two-col':
      shapes = compileTwoCol(s, theme)
      break
    case 'closing':
      shapes = compileClosing(s, theme)
      break
    default:
      shapes = compileBullets(s, theme)
  }
  if (s.diagram) {
    // Same content region compileBullets/bulletsBlock use — below the title
    // band when there is one, above the page-number chip's reserved strip.
    // See SlideSpec.diagram's doc comment (types.ts): additive on top of
    // whatever the slide's own kind produced, any kind may carry one.
    const top = s.title ? 2.3 : 1.1
    shapes = shapes.concat(compileDiagramShapes(s.diagram, theme, { x: 0.9, y: top, w: 11.4, h: Math.max(1, 6.6 - top) }))
  }
  if (s.kind !== 'title' && total > 1) shapes.push(pageNumberShape(index, total, theme))
  if (s.shapeOverrides) shapes = shapes.map((shape) => applyOverride(shape, s.shapeOverrides))
  if (s.shapeAnimations) shapes = shapes.map((shape) => applyAnimation(shape, s.shapeAnimations))
  const slide: WireSlide = { shapes, background: theme.background }
  if (s.transition) {
    slide.transition = s.transition.kind
    if (s.transition.direction) slide.transitionDirection = s.transition.direction
  }
  return slide
}

/** Compile a whole DeckSpec into the wire Deck format
 *  POST /v1/pptx/build expects. Pure — same input, same output. */
export function compileDeckToWire(deck: DeckSpec): WireDeck {
  const theme = resolveTheme(deck.theme)
  const total = deck.slides.length
  return {
    slides: deck.slides.map((s, i) => compileSlide(s, theme, i, total)),
  }
}
