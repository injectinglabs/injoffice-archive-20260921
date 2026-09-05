import { ChartImageExportManager } from '../../../packages/charts/src/imageExport'
import { toWireCharts } from '../../../packages/charts/src/toFile'
import type { ChartSpec } from '../../../packages/charts/src/types'

export function playgroundChartWire(spec: ChartSpec, grid: unknown[][]) {
  return toWireCharts([{ spec }], {
    sheetNameOf: (sheetId) => sheetId === spec.range.sheetId ? 'Forecast' : null,
    readRange: (range) => {
      const rows = grid.slice(range.startRow, range.endRow + 1)
      return rows.map((row) => (row ?? []).slice(range.startColumn, range.endColumn + 1))
    },
  })
}

export async function playgroundChartSvgExport(spec: ChartSpec, title: string) {
  const manager = new ChartImageExportManager(
    (id) => id === spec.id ? spec : undefined,
    {
      async render(request) {
        const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="320" height="180" viewBox="0 0 320 180"><rect width="320" height="180" fill="#f8fafc"/><text x="16" y="96" font-size="18">${escapeXml(request.chart.title ?? title)}</text></svg>`
        return {
          mediaType: 'image/svg+xml' as const,
          bytes: new TextEncoder().encode(svg),
          width: 320,
          height: 180,
        }
      },
    },
    { createJobId: () => `chart-export-${spec.id}` },
  )
  return manager.export(spec.id, { format: 'svg' })
}

function escapeXml(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;')
}
