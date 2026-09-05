// @vitest-environment node
// ^ same reason as DeckCanvasView.textOverflow.test.ts: react-konva's own
// import chain switches Konva into its browser code path under jsdom.

import { describe, expect, it } from 'vitest'
import 'konva/canvas-backend'
import Konva from 'konva'
import { shapeText } from './DeckCanvasView'
import type { WireShape } from './wire'

// Regression coverage for a live-staging bug (round 6 of the Slides OOXML
// pivot's S9): a REAL uploaded .pptx's shapes went blank in two ways once
// the earlier read-side fixes (blob URL recovery, the loading hang,
// placeholder geometry + p:style fill inheritance) were already in place —
// (1) an autoshape's OWN text ("Drag me" typed inside a rounded-rectangle
// callout) never painted, even though the shape's fill/stroke now rendered
// correctly, and (2) a title/body PLACEHOLDER (kind "" — a real .pptx
// placeholder never carries an explicit prstGeom of its own) rendered as
// literally nothing at all, text included.
//
// Root cause: KonvaShapeInner's entire text-rendering block was gated
// behind `shape.kind === 'textBox'` — every OTHER kind (or an unresolved
// "" kind) fell through to a branch that only ever paints an outline
// (fill/stroke), never text, REGARDLESS of whether shape.paragraphs had
// real content. This was invisible in every prior round because
// compile.ts's own textBox()/rect() helpers never produce a non-textBox
// shape with paragraphs attached — a DeckSpec-driven deck literally cannot
// exercise this path; only a REAL uploaded .pptx (where a shape commonly
// carries both an outline AND its own text, or a placeholder carries text
// with no resolved kind at all) can.
//
// Fix: text rendering (shapeText) is now called for every shape that HAS
// paragraphs, independent of its outline/kind — the two are siblings, not
// mutually exclusive branches. The property this suite proves is exactly
// that independence: shapeText's output depends ONLY on shape.paragraphs,
// never on shape.kind.

// A realistic EMU->px scale (a 960px-wide render of a standard 16:9 slide,
// cx=12192000 EMU) — NOT 1. shapeText runs a real shape's font size through
// ptToPx(pt, scale); at scale=1 a 14pt run becomes ptToPx(14,1) = 14*12700 =
// 177800 "px", so wrapLineToWidth wraps a two-word label onto two lines
// against ANY reasonably-sized box width — a geometry mistake in an early
// draft of this test, not a real bug (caught by actually running it: it
// failed the SAME way regardless of the fix, which is what said "this
// isn't testing what you think it's testing").
const SCALE = 960 / 12192000
const BOX_W = 2743200 * SCALE // the real roundRect's own width from the live repro
const BOX_H = 914400 * SCALE

function textShape(kind: WireShape['kind'], overrides: Partial<WireShape> = {}): WireShape {
  return {
    kind,
    x: 0,
    y: 0,
    cx: 2743200,
    cy: 914400,
    paragraphs: [{ runs: [{ text: 'Drag me' }] }],
    ...overrides,
  }
}

describe('shapeText — kind-independence (the actual bug)', () => {
  it('renders real text for a roundRect autoshape with its own paragraphs (an OOXML callout box) — this is what silently returned null before the fix', () => {
    const el = shapeText(textShape('roundRect'), BOX_W, BOX_H, SCALE)
    expect(el).not.toBeNull()
    expect(el?.props.text).toBe('Drag me')
  })

  it('renders real text for an unresolved/empty kind (a real .pptx title/body placeholder, which never carries an explicit prstGeom)', () => {
    // WireShapeKind's TS union doesn't include "", matching pptxpatch's own
    // permissive-on-read philosophy (an unresolved kind is surfaced as
    // itself, not rejected) — cast the same way a real parsed response
    // would arrive over the wire.
    const shape = textShape('' as WireShape['kind'], {
      placeholder: 'title',
      cx: 8229600,
      cy: 1143000,
      paragraphs: [{ runs: [{ text: 'Independently Built Test Deck' }] }],
    })
    const el = shapeText(shape, 8229600 * SCALE, 1143000 * SCALE, SCALE)
    expect(el).not.toBeNull()
    expect(el?.props.text).toBe('Independently Built Test Deck')
  })

  it('still renders for the original textBox kind (no regression on the case that always worked)', () => {
    const el = shapeText(textShape('textBox'), BOX_W, BOX_H, SCALE)
    expect(el).not.toBeNull()
    expect(el?.props.text).toBe('Drag me')
  })

  it('renders each of the other autoshape kinds too — text is independent of EVERY kind, not just roundRect', () => {
    for (const kind of ['rect', 'ellipse', 'triangle', 'diamond', 'rightArrow', 'pentagon', 'hexagon', 'star5'] as const) {
      const el = shapeText(textShape(kind), BOX_W, BOX_H, SCALE)
      expect(el?.props.text, `kind=${kind}`).toBe('Drag me')
    }
  })

  it('returns null for a shape with no paragraphs at all (an outline-only autoshape stays outline-only)', () => {
    expect(shapeText(textShape('roundRect', { paragraphs: [] }), BOX_W, BOX_H, SCALE)).toBeNull()
    expect(shapeText(textShape('roundRect', { paragraphs: undefined }), BOX_W, BOX_H, SCALE)).toBeNull()
  })
})

// A real-Konva pixel check that the RENDERING PATTERN this fix relies on
// (an outline node, then a Text node, as siblings painted in that order —
// see KonvaShapeInner) actually paints visible, distinguishable text over a
// filled shape, not just that shapeText computes the right props in
// isolation. Same node-canvas harness and same honest scope note as
// DeckCanvasView.textOverflow.test.ts: this proves the Konva integration
// for the mechanism the fix targets, not a literal replay of
// KonvaShapeInner's own branching (which isn't exported — see that
// function's own doc comment for why this suite tests shapeText directly
// instead).
describe('outline + text as siblings — real Konva render', () => {
  it('a filled rect with a text node painted after it shows BOTH the fill color and distinct darker text-colored pixels', () => {
    const w = Math.round(BOX_W)
    const h = Math.round(BOX_H)
    const stage = new Konva.Stage({ width: w, height: h })
    const layer = new Konva.Layer()
    stage.add(layer)

    const group = new Konva.Group({ x: 0, y: 0, width: w, height: h })
    layer.add(group)

    const fillHex = '#4f81bd' // the real resolved theme accent color from the live repro
    group.add(new Konva.Rect({ x: 0, y: 0, width: w, height: h, fill: fillHex }))
    const el = shapeText(textShape('roundRect', { paragraphs: [{ runs: [{ text: 'Drag me' }] }] }), w, h, SCALE)
    // el is a react-konva <Text> element; construct the equivalent real
    // Konva.Text node from the same props shapeText computed (react-konva
    // itself is what normally does this translation via its reconciler,
    // not exercised here — see the doc comment above).
    group.add(
      new Konva.Text({
        x: el!.props.x,
        y: el!.props.y,
        width: el!.props.width,
        height: el!.props.height,
        text: el!.props.text,
        fontSize: el!.props.fontSize,
        fontFamily: el!.props.fontFamily,
        fill: el!.props.fill,
        align: el!.props.align,
      }),
    )
    layer.draw()

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ctx = (layer.getCanvas().getContext() as any)._context as CanvasRenderingContext2D
    const fillRgb = [0x4f, 0x81, 0xbd]
    let sawFill = false
    let sawText = false
    for (let x = 0; x < w; x += 2) {
      for (let y = 0; y < h; y += 2) {
        const [r, g, b] = ctx.getImageData(x, y, 1, 1).data
        if (Math.abs(r - fillRgb[0]) < 10 && Math.abs(g - fillRgb[1]) < 10 && Math.abs(b - fillRgb[2]) < 10) sawFill = true
        // Text renders '#000000' by default — distinctly darker than the
        // medium-blue fill on every channel, not just low-contrast.
        if (r < 80 && g < 80 && b < 80) sawText = true
      }
    }
    expect(sawFill).toBe(true)
    expect(sawText).toBe(true)
  })
})
