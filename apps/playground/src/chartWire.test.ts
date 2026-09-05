import { describe, expect, it } from 'vitest'
import { playgroundChartSvgExport, playgroundChartWire } from './chartWire'
import type { ChartSpec } from '../../../packages/charts/src/types'

const spec: ChartSpec = {
  id: 'revenue-demo',
  type: 'column',
  title: 'Revenue vs target',
  range: { sheetId: 'forecast', startRow: 0, startColumn: 0, endRow: 2, endColumn: 2 },
  firstRowIsHeader: true,
  firstColumnIsCategory: true,
}

const grid = [
  ['Month', 'Revenue', 'Target'],
  ['Jan', 128, 135],
  ['Feb', 156, 145],
]

describe('playground chart native wire', () => {
  it('emits an OOXML-writable column chart and skips unsupported types', () => {
    const wired = playgroundChartWire(spec, grid)
    expect(wired.skipped).toEqual([])
    expect(wired.charts[0]).toMatchObject({ sheetName: 'Forecast', type: 'column', title: 'Revenue vs target' })
    expect(wired.charts[0]?.series.length).toBeGreaterThan(0)

    const skipped = playgroundChartWire({ ...spec, type: 'waterfall' }, grid)
    expect(skipped.charts).toEqual([])
    expect(skipped.skipped[0]).toMatch(/waterfall/)
  })

  it('exports an SVG artifact through the image-export manager', async () => {
    const artifact = await playgroundChartSvgExport(spec, 'Revenue vs target')
    expect(artifact.mediaType).toBe('image/svg+xml')
    expect(new TextDecoder().decode(artifact.bytes)).toContain('<svg')
    expect(artifact.width).toBe(320)
  })
})
