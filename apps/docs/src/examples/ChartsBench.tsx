import { useMemo, useState } from 'react'
import { buildEChartsOption, CHART_TYPES, EChartsPreview, extractChartData, type ChartType } from '@injoffice/charts'
import { LiveBench } from '../components/LiveBench'

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun']
const START = [128, 156, 149, 188, 214, 246]

export function ChartsBench() {
  const [type, setType] = useState<ChartType>('column')
  const [revenue, setRevenue] = useState(START)
  const spec = useMemo(() => ({
    id: 'docs-revenue',
    type,
    title: 'Revenue',
    range: { sheetId: 'forecast', startRow: 0, startColumn: 0, endRow: 6, endColumn: 1 },
    firstRowIsHeader: true,
    firstColumnIsCategory: true,
    valueFormat: 'currency' as const,
  }), [type])
  const grid = useMemo(() => [['Month', 'Revenue'], ...MONTHS.map((month, index) => [month, revenue[index]])], [revenue])
  const option = useMemo(() => buildEChartsOption(spec, extractChartData(grid, spec)), [grid, spec])
  return (
    <LiveBench title="Live example" hint="@injoffice/charts · extractChartData + buildEChartsOption">
      <div className="bench-controls">
        <label className="field">
          Chart type
          <select value={type} onChange={(event) => setType(event.target.value as ChartType)}>
            {CHART_TYPES.map((chartType) => <option key={chartType}>{chartType}</option>)}
          </select>
        </label>
        {MONTHS.map((month, index) => (
          <label className="field" key={month}>
            {month}
            <input type="number" value={revenue[index]} onChange={(event) => setRevenue((current) => current.map((value, at) => at === index ? Number(event.target.value) : value))} />
          </label>
        ))}
      </div>
      <EChartsPreview option={option} ariaLabel={`${type} chart of monthly revenue`} style={{ height: 280 }} />
    </LiveBench>
  )
}
