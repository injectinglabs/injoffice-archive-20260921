// Chart model for @injoffice/charts.
//
// The spec is deliberately independent of both Univer and ECharts: it is the
// unit that gets persisted (float-dom `data` → workbook snapshot), driven by
// agents (a chart.create RPC carries exactly this shape), and later mapped to
// OOXML DrawingML by the fidelity layer. Renderers and hosts come and go;
// the spec is the contract.

/** The chart names exposed by Univer's public Charts guide. These names are
 * renderer-neutral: an adapter normalizes them before producing ECharts (or a
 * future renderer) options. Native XLSX support is a separate capability. */
export const CANONICAL_CHART_TYPES = [
  'Line',
  'Column',
  'ColumnStacked',
  'ColumnPercentStacked',
  'Bar',
  'BarStacked',
  'BarPercentStacked',
  'Pie',
  'Donut',
  'Area',
  'AreaStacked',
  'AreaPercentStacked',
  'Radar',
  'Scatter',
  'Combination',
  'WordCloud',
  'Funnel',
  'Bubble',
  'Relation',
  'Waterfall',
  'Pareto',
  'Sankey',
  'Heatmap',
  'Boxplot',
  'Candlestick',
  'Histogram',
  'Treemap',
  'Sunburst',
  'Gauge',
  'Chord',
] as const

export type CanonicalChartType = (typeof CANONICAL_CHART_TYPES)[number]

/** Identifiers accepted by ChartSpec before the canonical public vocabulary
 * was introduced. They remain valid for saved snapshots and existing callers. */
export type LegacyChartType =
  | 'column'
  | 'bar'
  | 'line'
  | 'area'
  | 'pie'
  | 'doughnut'
  | 'scatter'
  | 'waterfall'
  | 'heatmap'
  | 'radar'
  | 'funnel'
  | 'treemap'
  | 'sankey'
  | 'boxplot'

export const LEGACY_CHART_TYPES: readonly LegacyChartType[] = [
  'column',
  'bar',
  'line',
  'area',
  'pie',
  'doughnut',
  'scatter',
  'waterfall',
  'heatmap',
  'radar',
  'funnel',
  'treemap',
  'sankey',
  'boxplot',
]

/** Every value accepted by ChartSpec. New UI should present
 * CANONICAL_CHART_TYPES; CHART_TYPES remains the discoverable public list. */
export type ChartType = CanonicalChartType | LegacyChartType
export const CHART_TYPES: readonly CanonicalChartType[] = CANONICAL_CHART_TYPES

const LEGACY_TO_CANONICAL: Readonly<Record<LegacyChartType, CanonicalChartType>> = {
  column: 'Column',
  bar: 'Bar',
  line: 'Line',
  area: 'Area',
  pie: 'Pie',
  doughnut: 'Donut',
  scatter: 'Scatter',
  waterfall: 'Waterfall',
  heatmap: 'Heatmap',
  radar: 'Radar',
  funnel: 'Funnel',
  treemap: 'Treemap',
  sankey: 'Sankey',
  boxplot: 'Boxplot',
}

const CANONICAL_SET: ReadonlySet<string> = new Set(CANONICAL_CHART_TYPES)
const LEGACY_SET: ReadonlySet<string> = new Set(LEGACY_CHART_TYPES)

/** Normalize persisted legacy names to the canonical renderer-neutral name. */
export function normalizeChartType(type: ChartType): CanonicalChartType {
  return LEGACY_SET.has(type) ? LEGACY_TO_CANONICAL[type as LegacyChartType] : (type as CanonicalChartType)
}

/** Runtime guard for untrusted JSON/RPC payloads. */
export function isChartType(type: unknown): type is ChartType {
  return typeof type === 'string' && (CANONICAL_SET.has(type) || LEGACY_SET.has(type))
}

/** The subset of types a per-series override may use (combo charts). */
export type SeriesComboType = 'column' | 'line' | 'area'

export type TrendlineKind = 'linear' | 'movingAverage'

/** Per-series presentation overrides, keyed by series name in ChartSpec.
 *  Only meaningful for the categorical family. */
export interface SeriesOverride {
  /** Render this series as a different mark than the chart's base type. */
  type?: SeriesComboType
  /** Bind the series to the secondary (right) value axis. */
  secondaryAxis?: boolean
  /** Overlay a computed trendline for this series. */
  trendline?: TrendlineKind
  /** Window for movingAverage (default 3). */
  trendlineWindow?: number
}

/** How axis/tooltip values render. 'auto' = plain numbers with thousand
 *  grouping; later phases feed this from the source cells' number formats. */
export type ValueFormat = 'auto' | 'percent' | 'currency' | 'plain'

/** Host-supplied theming; every field optional, sensible neutral defaults. */
export interface ChartTheme {
  palette?: string[]
  textColor?: string
  mutedTextColor?: string
  gridLineColor?: string
  background?: string
}

/** Back-to-front operations for charts on the same worksheet. Snapshot array
 * order is the durable model; a browser host can mirror it into its drawing
 * layer through ChartLayerHost. */
export type ChartLayerOperation = 'bringForward' | 'sendBackward' | 'bringToFront' | 'sendToBack'

/** A rectangular block of cells on one sheet, 0-indexed, inclusive. */
export interface CellRangeRef {
  sheetId: string
  startRow: number
  startColumn: number
  endRow: number
  endColumn: number
}

export interface ChartSpec {
  id: string
  /** Stable identity returned by native XLSX hydration. It is absent on
   * browser-created charts until the host saves and rehydrates. */
  readonly nativeIdentity?: NativeChartIdentity
  type: ChartType
  title?: string
  /** Source data block. First row is treated as series names and first column
   *  as categories when the heuristics in extract.ts say so (or when pinned
   *  via the flags below). */
  range: CellRangeRef
  /** Pin header interpretation instead of relying on the heuristic. */
  firstRowIsHeader?: boolean
  firstColumnIsCategory?: boolean
  /** Show a legend (default true when there is more than one series). */
  legend?: boolean
  /** Per-series overrides (combo type, secondary axis, trendlines), keyed by
   *  series name as extracted. Unknown names are ignored, so a stale override
   *  survives a rename harmlessly. */
  series?: Record<string, SeriesOverride>
  /** Value rendering for axes/tooltips. */
  valueFormat?: ValueFormat
  /** Currency symbol when valueFormat is 'currency' (default '$'). */
  currencySymbol?: string
}

/** One top-level chart in an XLSX worksheet drawing. Names, anchors, and
 * relationship ids are mutable; part plus drawing cNvPr id is the stable key. */
export interface NativeChartIdentity {
  readonly part: string
  readonly drawingPart: string
  readonly objectId: number
}

/** Extracted, renderer-agnostic data: what extract.ts produces from a values
 *  grid and what option.ts consumes. */
export interface ChartData {
  categories: string[]
  series: { name: string; values: (number | null)[] }[]
}
