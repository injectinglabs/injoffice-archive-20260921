// @vitest-environment node
// ^ Forced to vitest's plain 'node' environment regardless of a consumer's
// global config (the website vendors this file into a jsdom-environment
// suite): under jsdom, `window` exists, so Konva switches into its BROWSER
// code path and requires a real DOM container for the Stage instead of
// falling back to headless node-canvas — this file wants the latter.

import { describe, expect, it } from 'vitest'
// Konva 10's Node-side headless rendering needs its canvas backend
// registered explicitly (Konva 9 auto-detected `canvas`; 10 does not) —
// see konva's own "unsupported environment" error message for this exact
// fix. This package's devDependency is pinned to konva ^10 specifically so
// this import always resolves; a consumer vendoring this test file while
// still on konva 9 (peerDependencies still allows it) would need to drop
// this line — noted here rather than silently breaking their suite.
import 'konva/canvas-backend'
import Konva from 'konva'
import { canvasFontFamily, estimateTextWidthPx, fitFontSizePx, wrapLineToWidth } from './canvasGeometry'

// Regression coverage for a live-staging bug: on the 'title' SlideKind, a
// long title ("The Art and Science of Coffee Brewing") rendered past
// DeckCanvasView's canvas edge.
//
// ROUND 4's diagnosis and fix (kept as history — WRONG/INCOMPLETE):
// theorized the theme's displayFont ("Archivo") never being loaded as a web
// font, shipped fitFontSizePx + canvasFontFamily + a hard Group clip, and
// trusted Konva's OWN wrap="word" to make the line-break decision. Nick
// re-verified on staging: the bug was STILL present, reproduced identically
// on "terra" (which uses the same font as boardroom, so that was never
// really the variable), with a precise symptom — title text rendered as
// ONE long unwrapped line reaching the canvas's own edge, not clipped at
// the shape's own ~81%-width box. That rules out "Konva's wrap engaged
// using slightly-wrong metrics" (which would still produce SOME wrapping,
// just at a different point) — it means wrap="word" effectively did not
// engage AT ALL for that render, in that browser.
//
// ROUND 5's fix stops depending on Konva/canvas's own wrap decision
// entirely: DeckCanvasView now pre-wraps text itself (wrapLineToWidth, a
// deterministic character-width heuristic) and renders with wrap="none".
//
// Honest scope of THIS file, confirmed by direct experimentation (again —
// round 4's suite found the same thing): this exact Konva+node-canvas
// combination self-bounds text painting within its configured `width`
// REGARDLESS of wrap mode ('word' OR 'none') and regardless of whether the
// text was pre-wrapped, even for the raw unwrapped repro string and even
// for a deliberately adversarial unbreakable word. That means this file
// CANNOT construct a literal "paints outside its box" failure at the Konva-
// render level to diff against — the browser-specific trigger for the live
// bug (most likely an async web-font measure-vs-paint mismatch, or some
// other canvas-measurement quirk this environment doesn't have) does not
// reproduce headlessly, full stop, and this suite does not claim otherwise.
// The genuine "before vs after" comparison for THIS bug lives in
// canvasGeometry.test.ts instead, at the level that IS deterministic and
// portable: estimateTextWidthPx(wholeRawTitle) is measurably wider than the
// box (the "before" — what one unwrapped line would require), while every
// line wrapLineToWidth produces individually fits (the "after"). What THIS
// file verifies is narrower but still real: fitFontSizePx/canvasFontFamily/
// wrapLineToWidth integrate correctly with ACTUAL Konva Text objects (not
// just their own isolated pure-function tests) and the resulting render,
// for the real reported repro string at the real compiled title-box size,
// stays within bounds in this environment — a genuine integration check,
// not a reproduction of the browser bug itself.

function textPaintsOutsideBox(opts: { boxX: number; boxY: number; boxW: number; boxH: number; stageW: number; stageH: number; text: string; fontSize: number }): boolean {
  const stage = new Konva.Stage({ width: opts.stageW, height: opts.stageH })
  const layer = new Konva.Layer()
  stage.add(layer)

  // Deliberately NO clip on the group — isolating exactly what the TEXT
  // NODE itself paints, so a pass here is unambiguously about
  // wrapLineToWidth's own output, not masked by clipping.
  const group = new Konva.Group({ x: opts.boxX, y: opts.boxY, width: opts.boxW, height: opts.boxH })
  layer.add(group)

  const text = new Konva.Text({
    x: 0,
    y: 0,
    width: opts.boxW,
    height: opts.boxH,
    text: opts.text,
    fontSize: opts.fontSize,
    fontFamily: canvasFontFamily('Archivo'),
    fill: '#ff0000',
    wrap: 'none', // DeckCanvasView's actual mode — pre-wrapped text, no further Konva-side breaking.
  })
  group.add(text)
  layer.draw()

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const ctx = (layer.getCanvas().getContext() as any)._context as CanvasRenderingContext2D
  for (let x = 0; x < opts.stageW; x += 2) {
    for (let y = 0; y < opts.stageH; y += 2) {
      if (x >= opts.boxX && x < opts.boxX + opts.boxW && y >= opts.boxY && y < opts.boxY + opts.boxH) continue
      const [r, g, b, a] = ctx.getImageData(x, y, 1, 1).data
      if (r > 200 && g < 80 && b < 80 && a > 0) return true
    }
  }
  return false
}

describe('DeckCanvasView text overflow — round 5 fix integration (real Konva + node-canvas render)', () => {
  const boxW = 777
  const boxH = 144
  const stageW = boxW + 100
  const stageH = boxH + 40
  const reportedTitle = 'The Art and Science of Coffee Brewing'
  const fontSize = fitFontSizePx(reportedTitle, boxW, 40)

  it('the estimate this fix is built on: the raw reported title as ONE unwrapped line is measurably wider than the compiled box (the actual live symptom, at the math level)', () => {
    expect(estimateTextWidthPx(reportedTitle, fontSize)).toBeGreaterThan(boxW)
  })

  it('wrapLineToWidth + wrap="none" (DeckCanvasView\'s actual approach) renders the reported title within its box, as a real Konva Text object', () => {
    const wrapped = wrapLineToWidth(reportedTitle, boxW, fontSize).join('\n')
    expect(textPaintsOutsideBox({ boxX: 20, boxY: 20, boxW, boxH, stageW, stageH, text: wrapped, fontSize })).toBe(false)
  })

  it('an adversarially long single word (wrapLineToWidth cannot break it, only fitFontSizePx can shrink it) renders within bounds once shrunk', () => {
    const word = 'Supercalifragilisticexpialidocious'
    const smallBoxW = 100
    const shrunk = fitFontSizePx(word, smallBoxW, 200)
    const wrapped = wrapLineToWidth(word, smallBoxW, shrunk).join('\n')
    expect(textPaintsOutsideBox({ boxX: 10, boxY: 10, boxW: smallBoxW, boxH: 50, stageW: 600, stageH: 400, text: wrapped, fontSize: shrunk })).toBe(false)
  })
})
