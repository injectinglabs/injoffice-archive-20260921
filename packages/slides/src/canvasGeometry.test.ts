import { describe, expect, it } from 'vitest'
import {
  EMU_PER_INCH,
  EMU_PER_PT,
  MIN_SHAPE_PX,
  canvasFontFamily,
  dragResultEmu,
  emuToPx,
  estimateTextWidthPx,
  fitFontSizePx,
  makeViewport,
  paragraphLine,
  ptToPx,
  regularPolygonSides,
  resizeResultEmu,
  roundRectCornerRadiusPx,
  shapeRectPx,
  wrapLineToWidth,
} from './canvasGeometry'

describe('makeViewport', () => {
  it('scales the 16:9 default slide size to a requested width, height follows the ratio', () => {
    const vp = makeViewport(12192000, 6858000, 960)
    expect(vp.widthPx).toBe(960)
    expect(vp.heightPx).toBeCloseTo(540, 5) // 960 * 9/16
    expect(vp.scale).toBeCloseTo(960 / 12192000, 10)
  })

  it('is degenerate-safe for a zero slide width', () => {
    const vp = makeViewport(0, 6858000, 960)
    expect(vp.scale).toBe(0)
  })
})

describe('emuToPx / ptToPx', () => {
  it('emuToPx is a straight linear scale', () => {
    expect(emuToPx(914400, 0.001)).toBeCloseTo(914.4, 5)
  })

  it('ptToPx goes through EMU at 12700 EMU/pt', () => {
    // 1pt at scale 1 (1px per EMU) = 12700px — deliberately absurd scale, just checks the constant.
    expect(ptToPx(1, 1)).toBe(12700)
    expect(ptToPx(0, 5)).toBe(0)
  })
})

describe('shapeRectPx', () => {
  it('converts all four fields at the same scale', () => {
    const r = shapeRectPx({ x: 100, y: 200, cx: 300, cy: 400 }, 0.01)
    expect(r).toEqual({ x: 1, y: 2, w: 3, h: 4 })
  })
})

describe('regularPolygonSides', () => {
  it('maps the polygon kinds to their side count', () => {
    expect(regularPolygonSides('triangle')).toBe(3)
    expect(regularPolygonSides('diamond')).toBe(4)
    expect(regularPolygonSides('pentagon')).toBe(5)
    expect(regularPolygonSides('hexagon')).toBe(6)
  })

  it('is null for non-polygon kinds', () => {
    expect(regularPolygonSides('ellipse')).toBeNull()
    expect(regularPolygonSides('rect')).toBeNull()
    expect(regularPolygonSides('star5')).toBeNull()
    expect(regularPolygonSides('rightArrow')).toBeNull()
    expect(regularPolygonSides('line')).toBeNull()
    expect(regularPolygonSides('textBox')).toBeNull()
  })
})

describe('roundRectCornerRadiusPx', () => {
  it('is 16% of the shorter side', () => {
    expect(roundRectCornerRadiusPx(100, 50)).toBeCloseTo(8, 5)
    expect(roundRectCornerRadiusPx(50, 100)).toBeCloseTo(8, 5)
  })
})

describe('dragResultEmu', () => {
  it('converts a new top-left px position back to EMU at the given scale (round-trip with shapeRectPx)', () => {
    const scale = 960 / 12192000
    const original = { x: 838200, y: 365125, cx: 100, cy: 100 }
    const r = shapeRectPx(original, scale)
    // Simulate the drag landing back exactly where it started.
    expect(dragResultEmu(r.x, r.y, scale)).toEqual({ x: original.x, y: original.y })
  })

  it('is degenerate-safe for a zero scale', () => {
    expect(dragResultEmu(100, 100, 0)).toEqual({ x: 0, y: 0 })
  })

  it('moves right/down for positive new coordinates', () => {
    const scale = 0.001
    expect(dragResultEmu(2000, 3000, scale)).toEqual({ x: 2000000, y: 3000000 })
  })
})

describe('resizeResultEmu', () => {
  it('converts a new px size back to EMU at the given scale', () => {
    const scale = 0.001
    expect(resizeResultEmu(3000, 4000, scale)).toEqual({ cx: 3000000, cy: 4000000 })
  })

  it('floors at MIN_SHAPE_PX so a shape can never collapse to zero/negative size', () => {
    const scale = 0.001
    const r = resizeResultEmu(1, 1, scale)
    expect(r.cx).toBe(Math.round(MIN_SHAPE_PX / scale))
    expect(r.cy).toBe(Math.round(MIN_SHAPE_PX / scale))
    const negative = resizeResultEmu(-50, -50, scale)
    expect(negative.cx).toBe(Math.round(MIN_SHAPE_PX / scale))
    expect(negative.cy).toBe(Math.round(MIN_SHAPE_PX / scale))
  })

  it('is degenerate-safe for a zero scale', () => {
    expect(resizeResultEmu(100, 100, 0)).toEqual({ cx: 0, cy: 0 })
  })
})

describe('paragraphLine', () => {
  it('passes plain text through unchanged when not a bullet', () => {
    expect(paragraphLine('Hello', false, 0)).toBe('Hello')
    expect(paragraphLine('Hello', undefined, undefined)).toBe('Hello')
  })

  it('prefixes a bullet glyph at level 0', () => {
    expect(paragraphLine('Point one', true, 0)).toBe('• Point one')
    expect(paragraphLine('Point one', true, undefined)).toBe('• Point one')
  })

  it('indents and alternates the glyph by level', () => {
    expect(paragraphLine('Nested', true, 1)).toBe('  ◦ Nested')
    expect(paragraphLine('Deeper', true, 2)).toBe('    • Deeper')
  })
})

// ---- Text overflow defense (regression coverage for the live-staging bug:
// a long title on the 'title' SlideKind rendered past the canvas's right
// edge — see canvasGeometry.ts's own "Text overflow defense" comment for
// the full root-cause writeup and why it doesn't reproduce headlessly). ----

describe('estimateTextWidthPx', () => {
  it('scales linearly with both text length and font size', () => {
    expect(estimateTextWidthPx('', 40)).toBe(0)
    expect(estimateTextWidthPx('AAAA', 40)).toBeCloseTo(estimateTextWidthPx('AA', 40) * 2, 5)
    expect(estimateTextWidthPx('AAAA', 80)).toBeCloseTo(estimateTextWidthPx('AAAA', 40) * 2, 5)
  })
})

describe('fitFontSizePx', () => {
  it('leaves the font size untouched when the longest word already fits', () => {
    // The actual reported repro string — none of its words are
    // individually too long for a realistic title box, so this asserts
    // the fix does NOT (mis)shrink normal titles.
    const text = 'The Art and Science of Coffee Brewing'
    expect(fitFontSizePx(text, 777, 40)).toBe(40)
  })

  it('shrinks a font size a single long word could not otherwise wrap to fit', () => {
    const size = fitFontSizePx('Supercalifragilisticexpialidocious', 200, 40)
    expect(size).toBeLessThan(40)
    expect(estimateTextWidthPx('Supercalifragilisticexpialidocious', size)).toBeLessThanOrEqual(200)
  })

  it('never shrinks below minFontSizePx even for a pathologically long word', () => {
    const size = fitFontSizePx('X'.repeat(500), 50, 40, 12)
    expect(size).toBe(12)
  })

  it('is a no-op for empty text or a non-positive box width', () => {
    expect(fitFontSizePx('', 500, 40)).toBe(40)
    expect(fitFontSizePx('hello', 0, 40)).toBe(40)
    expect(fitFontSizePx('hello', -10, 40)).toBe(40)
  })
})

describe('canvasFontFamily', () => {
  it('appends a generic sans-serif fallback to a named font', () => {
    expect(canvasFontFamily('Archivo')).toBe('Archivo, sans-serif')
  })

  it('trims whitespace', () => {
    expect(canvasFontFamily('  Archivo  ')).toBe('Archivo, sans-serif')
  })

  it('is undefined for empty/whitespace-only/undefined input (Konva default)', () => {
    expect(canvasFontFamily(undefined)).toBeUndefined()
    expect(canvasFontFamily('')).toBeUndefined()
    expect(canvasFontFamily('   ')).toBeUndefined()
  })
})

describe('wrapLineToWidth', () => {
  it('does not wrap when the whole line already fits', () => {
    expect(wrapLineToWidth('short line', 1000, 20)).toEqual(['short line'])
  })

  it('greedily wraps a long line into multiple lines, each estimated to fit', () => {
    const line = 'The Art and Science of Coffee Brewing'
    const boxWidthPx = 200
    const fontSizePx = 20
    const lines = wrapLineToWidth(line, boxWidthPx, fontSizePx)
    expect(lines.length).toBeGreaterThan(1)
    for (const l of lines) {
      expect(estimateTextWidthPx(l, fontSizePx)).toBeLessThanOrEqual(boxWidthPx)
    }
    // Word order/content is preserved (internal spacing normalized to single spaces).
    expect(lines.join(' ').replace(/\s+/g, ' ')).toBe(line)
  })

  it('preserves a leading indent (bullet glyph) on every wrapped line', () => {
    const line = '  ◦ A fairly long nested bullet point that needs to wrap onto more than one line'
    const lines = wrapLineToWidth(line, 150, 16)
    expect(lines.length).toBeGreaterThan(1)
    for (const l of lines) {
      expect(l.startsWith('  ')).toBe(true)
    }
    expect(lines[0]?.trim().startsWith('◦')).toBe(true)
  })

  it('puts a single word wider than the box on its own (still overflowing) line rather than dropping it', () => {
    const lines = wrapLineToWidth('Supercalifragilisticexpialidocious', 50, 40)
    expect(lines).toEqual(['Supercalifragilisticexpialidocious'])
  })

  it('is a no-op for blank text or a degenerate box/font size', () => {
    expect(wrapLineToWidth('', 500, 20)).toEqual([''])
    expect(wrapLineToWidth('   ', 500, 20)).toEqual(['   '])
    expect(wrapLineToWidth('hello world', 0, 20)).toEqual(['hello world'])
    expect(wrapLineToWidth('hello world', 500, 0)).toEqual(['hello world'])
  })
})

describe('title-slide overflow regression (the actual reported repro)', () => {
  // Round 4's fix (fitFontSizePx alone, trusting Konva's own wrap="word")
  // did NOT hold on real staging — Nick reproduced the identical bug on
  // the "terra" theme (which uses the same font as boardroom, so "which
  // theme" was never the real variable) with a very specific symptom: the
  // title rendered as ONE long unwrapped line reaching the canvas's own
  // edge, not clipped at the shape's own ~81%-width box. That is
  // consistent with Konva's OWN width-based wrap decision not being
  // trustworthy in that browser, in a way this package's headless
  // node-canvas environment does not reproduce (confirmed while building
  // the round-4 fix). Round 5's fix stops depending on it: these tests
  // cover the manual pre-wrap (wrapLineToWidth) DeckCanvasView now uses
  // instead of Konva's wrap="word", at the ACTUAL compile.ts numbers.
  const scale = 960 / (13.333 * EMU_PER_INCH)
  const boxWidthPx = 10.8 * EMU_PER_INCH * scale
  const fontSizePx = 40 * EMU_PER_PT * scale
  const reportedTitle = 'The Art and Science of Coffee Brewing'

  it('the compiled title box is wide enough that the reported title needs no font shrink, at a realistic canvas width', () => {
    expect(fitFontSizePx(reportedTitle, boxWidthPx, fontSizePx)).toBe(fontSizePx)
  })

  it('the reported title, pre-wrapped at the actual compiled box/font size, produces lines that all fit the box — the direct fix for the live symptom (one long unwrapped line)', () => {
    const lines = wrapLineToWidth(reportedTitle, boxWidthPx, fontSizePx)
    expect(lines.length).toBeGreaterThan(1) // it DOES need to wrap — that's the whole point
    for (const l of lines) {
      expect(estimateTextWidthPx(l, fontSizePx)).toBeLessThanOrEqual(boxWidthPx)
    }
  })

  it('reproduces the reported failure mode directly: the SAME title, naively measured as one unwrapped line, is far wider than the box (this is what Konva rendered live when its own wrap did not engage)', () => {
    const unwrappedWidth = estimateTextWidthPx(reportedTitle, fontSizePx)
    expect(unwrappedWidth).toBeGreaterThan(boxWidthPx)
  })
})
