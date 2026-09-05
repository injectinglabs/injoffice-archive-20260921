import { useMemo, useState } from 'react'
import {
  buildEChartsOption,
  CHART_TYPES,
  EChartsPreview,
  extractChartData,
  fiveNumberSummary,
  linearTrend,
  movingAverage,
  type ChartType,
} from '@injoffice/charts'
import { playgroundChartSvgExport, playgroundChartWire } from '../chartWire'

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun']
const START_REVENUE = [128, 156, 149, 188, 214, 246]
const START_TARGET = [135, 145, 160, 175, 195, 220]
const START_LOW = [118, 140, 136, 170, 190, 208]
const START_HIGH = [142, 166, 171, 198, 226, 258]
const LINK_CATEGORIES = [
  'Pipeline → Qualified',
  'Qualified → Proposed',
  'Proposed → Won',
  'Pipeline → Nurture',
  'Nurture → Qualified',
  'Proposed → Lost',
]
const HIERARCHY_CATEGORIES = [
  'Revenue/North/Jan',
  'Revenue/North/Feb',
  'Revenue/South/Mar',
  'Revenue/South/Apr',
  'Revenue/West/May',
  'Revenue/West/Jun',
]
const LINK_TYPES: ReadonlySet<ChartType> = new Set(['Relation', 'Sankey', 'Chord'])
const TREND_TYPES: ReadonlySet<ChartType> = new Set(['Column', 'Line', 'Area', 'Combination'])

export default function ChartsPage() {
  const [type, setType] = useState<ChartType>('Column')
  const [revenue, setRevenue] = useState(START_REVENUE)
  const [showTrend, setShowTrend] = useState(true)
  const categories = LINK_TYPES.has(type) ? LINK_CATEGORIES : type === 'Sunburst' ? HIERARCHY_CATEGORIES : MONTHS
  const categoryHeader = LINK_TYPES.has(type) ? 'Flow' : type === 'Sunburst' ? 'Path' : 'Month'
  const supportsTrend = TREND_TYPES.has(type)
  const chartTitle = LINK_TYPES.has(type) ? 'Pipeline flow' : type === 'Sunburst' ? 'Revenue hierarchy' : 'Revenue vs target'

  const grid = useMemo(
    () => [
      [categoryHeader, 'Revenue', 'Target', 'Low', 'High'],
      ...categories.map((category, index) => [category, revenue[index], START_TARGET[index], START_LOW[index], START_HIGH[index]]),
    ],
    [categories, categoryHeader, revenue],
  )
  const spec = useMemo(
    () => ({
      id: 'revenue-demo',
      type,
      title: chartTitle,
      range: { sheetId: 'forecast', startRow: 0, startColumn: 0, endRow: 6, endColumn: 4 },
      firstRowIsHeader: true,
      firstColumnIsCategory: true,
      valueFormat: 'currency' as const,
      series: showTrend && supportsTrend ? { Revenue: { trendline: 'linear' as const } } : undefined,
    }),
    [chartTitle, showTrend, supportsTrend, type],
  )
  const data = useMemo(() => extractChartData(grid, spec), [grid, spec])
  const option = useMemo(() => buildEChartsOption(spec, data), [data, spec])
  const trend = useMemo(() => linearTrend(revenue), [revenue])
  const moving = useMemo(() => movingAverage(revenue, 3), [revenue])
  const summary = useMemo(() => fiveNumberSummary(revenue), [revenue])
  const wire = useMemo(() => playgroundChartWire(spec, grid), [grid, spec])
  const [exportNote, setExportNote] = useState('No image export yet.')
  return (
    <section className="tool-page" data-demo-surface="charts" aria-label="Chart analysis workbench">
      <div className="tool-page__controls" role="toolbar" aria-label="Chart controls">
        <label className="tool-field">
          Chart type
          <select value={type} onChange={(event) => setType(event.target.value as ChartType)}>
            {CHART_TYPES.map((chartType) => <option key={chartType}>{chartType}</option>)}
          </select>
        </label>
        <label className="tool-check">
          <input type="checkbox" checked={showTrend} disabled={!supportsTrend} onChange={(event) => setShowTrend(event.target.checked)} />
          Linear trendline (line, column, area, combination)
        </label>
        <button className="workbench-button" type="button" onClick={() => setRevenue(START_REVENUE)}>Reset data</button>
        <button className="workbench-button" type="button" onClick={() => {
          void playgroundChartSvgExport(spec, spec.title ?? type).then((artifact) => {
            setExportNote(`Exported ${artifact.mediaType} · ${artifact.bytes.byteLength} bytes`)
          }).catch((reason: unknown) => setExportNote(reason instanceof Error ? reason.message : String(reason)))
        }}>Export SVG</button>
        <span className="tool-page__status" role="status">{CHART_TYPES.length} chart types · ChartSpec to ECharts option and native wire</span>
      </div>

      <div className="tool-page__grid tool-page__grid--wide">
        <section className="tool-card tool-card--hero" aria-labelledby="chart-preview-title">
          <div className="tool-card__heading">
            <div><span className="tool-eyebrow">Interactive proof</span><h2 id="chart-preview-title">{chartTitle}</h2></div>
            <span className="tool-chip">{type}</span>
          </div>
          <EChartsPreview
            className="chart-proof chart-proof--echarts"
            option={option}
            ariaLabel={`${type} chart preview: ${chartTitle}`}
          />
        </section>

        <section className="tool-card" aria-labelledby="chart-data-title">
          <div className="tool-card__heading"><div><span className="tool-eyebrow">Source range</span><h2 id="chart-data-title">Edit the series</h2></div></div>
          <div className="tool-data-list">
            {categories.map((category, index) => (
              <label key={category} className="tool-data-row">
                <span>{category}</span>
                <input
                  aria-label={`${category} revenue`}
                  type="number"
                  min="0"
                  value={revenue[index]}
                  onChange={(event) => setRevenue((current) => current.map((value, at) => at === index ? Number(event.target.value) : value))}
                />
              </label>
            ))}
          </div>
        </section>

        <section className="tool-card" aria-labelledby="chart-analysis-title">
          <div className="tool-card__heading"><div><span className="tool-eyebrow">Analysis APIs</span><h2 id="chart-analysis-title">Computed from the same series</h2></div></div>
          <dl className="tool-metrics">
            <div><dt>Minimum</dt><dd>{summary?.[0] ?? '—'}</dd></div>
            <div><dt>Median</dt><dd>{summary?.[2] ?? '—'}</dd></div>
            <div><dt>Maximum</dt><dd>{summary?.[4] ?? '—'}</dd></div>
            <div><dt>Latest trend</dt><dd>{Math.round(trend.at(-1) ?? 0)}</dd></div>
            <div><dt>3-month average</dt><dd>{Math.round(moving.at(-1) ?? 0)}</dd></div>
          </dl>
        </section>

        <section className="tool-card tool-card--code" aria-labelledby="chart-option-title">
          <div className="tool-card__heading"><div><span className="tool-eyebrow">Renderer-neutral output</span><h2 id="chart-option-title">Generated option</h2></div></div>
          <pre>{JSON.stringify(option, null, 2)}</pre>
        </section>
        <section className="tool-card tool-card--code" aria-labelledby="chart-wire-title">
          <div className="tool-card__heading"><div><span className="tool-eyebrow">Native XLSX wire</span><h2 id="chart-wire-title">toWireCharts</h2></div></div>
          <pre>{JSON.stringify(wire.charts[0] ?? { skipped: wire.skipped }, null, 2)}</pre>
          <p className="tool-note">{exportNote}</p>
        </section>
      </div>
    </section>
  )
}
