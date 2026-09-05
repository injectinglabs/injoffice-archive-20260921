import { describe, expect, it } from 'vitest'
import {
  cellIsVisible,
  cssPixels,
  mapCellRectToOverlay,
  measureUniverCellAnchor,
  scrollForCell,
  type RectLike,
  type UniverCanvasLike,
} from './collabCellAnchor'

const rect = (left: number, top: number, width: number, height: number): RectLike => ({
  left,
  top,
  right: left + width,
  bottom: top + height,
  width,
  height,
})

describe('collaboration cell anchoring', () => {
  it('maps skeleton coordinates through scroll, zoom, canvas scaling, and overlay offset', () => {
    expect(mapCellRectToOverlay({
      cellRect: rect(210, 70, 100, 24),
      canvasRect: rect(300, 200, 600, 400),
      overlayRect: { left: 280, top: 150 },
      canvasCssWidth: 1200,
      canvasCssHeight: 800,
      zoom: 1.5,
      viewportScrollX: 50,
      viewportScrollY: 20,
    })).toEqual({
      left: 140,
      right: 215,
      top: 87.5,
      bottom: 105.5,
      width: 75,
      height: 18,
    })
  })

  it('preserves header offsets at zero scroll', () => {
    expect(mapCellRectToOverlay({
      cellRect: rect(46, 28, 80, 20),
      canvasRect: rect(100, 50, 500, 300),
      overlayRect: { left: 100, top: 50 },
      zoom: 1,
      viewportScrollX: 0,
      viewportScrollY: 0,
    })).toMatchObject({ left: 46, top: 28, width: 80, height: 20 })
  })

  it('does not scroll a cell on either frozen axis', () => {
    const scroll = { viewportScrollX: 90, viewportScrollY: 70 }
    const freeze = { startColumn: 2, startRow: 3 }
    expect(scrollForCell(scroll, freeze, 1, 1)).toEqual({ viewportScrollX: 0, viewportScrollY: 0 })
    expect(scrollForCell(scroll, freeze, 1, 4)).toEqual({ viewportScrollX: 90, viewportScrollY: 0 })
    expect(scrollForCell(scroll, freeze, 6, 1)).toEqual({ viewportScrollX: 0, viewportScrollY: 70 })
    expect(scrollForCell(scroll, freeze, 6, 4)).toEqual(scroll)
  })

  it('checks all visible viewports, including frozen panes', () => {
    const ranges = [
      { startRow: 0, endRow: 1, startColumn: 0, endColumn: 8 },
      { startRow: 20, endRow: 40, startColumn: 4, endColumn: 12 },
    ]
    expect(cellIsVisible(0, 7, ranges)).toBe(true)
    expect(cellIsVisible(25, 8, ranges)).toBe(true)
    expect(cellIsVisible(10, 8, ranges)).toBe(false)
  })

  it('scopes the render-canvas lookup to the supplied editor', () => {
    const canvas: UniverCanvasLike = {
      style: { width: '400px', height: '240px' },
      getBoundingClientRect: () => rect(120, 80, 400, 240),
    }
    const selectors: string[] = []
    const result = measureUniverCellAnchor({
      editorRoot: {
        querySelector(selector) {
          selectors.push(selector)
          return canvas
        },
      },
      overlayRoot: { getBoundingClientRect: () => rect(100, 60, 440, 280) },
      worksheet: {
        getRange: () => ({ getCellRect: () => rect(46, 28, 80, 20) }),
        getZoom: () => 1,
      },
      row: 0,
      column: 0,
      scroll: { viewportScrollX: 0, viewportScrollY: 0 },
    })
    expect(selectors).toEqual(['canvas[data-u-comp="render-canvas"]'])
    expect(result).toMatchObject({ left: 66, top: 48, width: 80, height: 20 })
  })

  it('fails closed when the cell is outside every viewport or the canvas is absent', () => {
    const worksheet = {
      getRange: () => ({ getCellRect: () => rect(46, 28, 80, 20) }),
      getZoom: () => 1,
      getVisibleRangesOfAllViewports: () => new Map([['main', { startRow: 4, endRow: 10, startColumn: 4, endColumn: 10 }]]),
    }
    expect(measureUniverCellAnchor({
      editorRoot: { querySelector: () => null },
      overlayRoot: { getBoundingClientRect: () => rect(0, 0, 100, 100) },
      worksheet,
      row: 0,
      column: 0,
      scroll: { viewportScrollX: 0, viewportScrollY: 0 },
    })).toBeNull()
  })

  it('parses positive CSS pixels and ignores invalid dimensions', () => {
    expect(cssPixels('810px')).toBe(810)
    expect(cssPixels('810.5px')).toBe(810.5)
    expect(cssPixels('auto')).toBeUndefined()
    expect(cssPixels('0px')).toBeUndefined()
  })
})
