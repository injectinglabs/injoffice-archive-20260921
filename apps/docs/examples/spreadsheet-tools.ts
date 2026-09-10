import { buildEChartsOption, extractChartData, type ChartSpec } from '@injoffice/charts'
import { bakePivot } from '@injoffice/pivots'

export function revenueChart() {
  const values = [['Month', 'Revenue'], ['July', 120], ['August', 165]]
  const spec: ChartSpec = {
    id: 'revenue', type: 'Column', title: 'Revenue',
    range: { sheetId: 'revenue-sheet', startRow: 0, startColumn: 0, endRow: 2, endColumn: 1 },
  }
  return buildEChartsOption(spec, extractChartData(values, spec))
}

export function salesByRegion() {
  return bakePivot(
    [['Region', 'Sales'], ['West', 20], ['East', 35], ['West', 15]],
    { rows: ['Region'], columns: [], values: [{ field: 'Sales', agg: 'sum' }] },
  )
}
