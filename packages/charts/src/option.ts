import { fiveNumberSummary, linearTrend, movingAverage, waterfallSegments } from './analysis'
import type { ChartData, ChartSpec, ChartTheme } from './types'
import { normalizeChartType } from './types'

// Pure mapping: ChartSpec + ChartData (+ optional ChartTheme) → an ECharts
// option object. Pure and dependency-free (the option is plain JSON) so it
// unit-tests without a DOM or an echarts instance. The renderer just calls
// setOption on the result.
//
// Data interpretation per specialty type (documented here because it IS the
// contract users experience):
//  - pie/donut/funnel/treemap/gauge: first series = values
//  - waterfall: first series = signed deltas per category; a computed Total
//    bar is appended
//  - heatmap: the whole numeric block is a matrix — series names on X,
//    categories on Y, cell value = intensity
//  - radar: each series is a polygon over the categories as indicators
//  - bubble: first three series = x, y, size
//  - candlestick: first four series = open, close, low, high
//  - histogram: first series = raw observations; bins are computed here
//  - sunburst: categories are slash-delimited hierarchy paths
//  - relation/sankey/chord: categories are "Source → Target" labels (also
//    accepts "->" or ">") and series[0] holds edge values
//  - boxplot: each series becomes one box from its five-number summary

const DEFAULT_PALETTE = ['#2f7ed8', '#0fa98f', '#f2a13c', '#c0453b', '#8461c9', '#5b8a72', '#c85f99', '#6b7379']

interface ResolvedTheme {
  palette: string[]
  textColor: string
  mutedTextColor: string
  gridLineColor: string
  background: string
}

function resolveTheme(theme?: ChartTheme): ResolvedTheme {
  return {
    palette: theme?.palette?.length ? theme.palette : DEFAULT_PALETTE,
    textColor: theme?.textColor ?? '#1d2427',
    mutedTextColor: theme?.mutedTextColor ?? '#6b7379',
    gridLineColor: theme?.gridLineColor ?? 'rgba(110, 125, 130, 0.25)',
    background: theme?.background ?? 'transparent',
  }
}

function valueFormatter(spec: ChartSpec): ((v: number) => string) | undefined {
  const sym = spec.currencySymbol ?? '$'
  switch (spec.valueFormat) {
    case 'percent':
      return (v: number) => `${Math.round(v * 1000) / 10}%`
    case 'currency':
      return (v: number) => `${sym}${Number(v).toLocaleString()}`
    case 'plain':
      return (v: number) => String(v)
    case 'auto':
    default:
      return (v: number) => Number(v).toLocaleString()
  }
}

function axisCommon(t: ResolvedTheme) {
  return {
    axisLabel: { color: t.mutedTextColor },
    axisLine: { lineStyle: { color: t.gridLineColor } },
    splitLine: { lineStyle: { color: t.gridLineColor } },
  }
}

function percentSeries(data: ChartData): ChartData['series'] {
  return data.series.map((series) => ({
    ...series,
    values: series.values.map((value, index) => {
      if (value === null) return null
      const total = data.series.reduce((sum, candidate) => sum + Math.abs(candidate.values[index] ?? 0), 0)
      return total === 0 ? 0 : value / total
    }),
  }))
}

function linkData(data: ChartData): {
  nodes: { name: string }[]
  links: { source: string; target: string; value: number }[]
} {
  const links: { source: string; target: string; value: number }[] = []
  const nodes = new Set<string>()
  data.categories.forEach((label, index) => {
    const match = /^(.+?)\s*(?:→|->|>)\s*(.+)$/.exec(label)
    const value = data.series[0]?.values[index]
    if (!match || value === null || value === undefined) return
    const source = match[1].trim()
    const target = match[2].trim()
    if (!source || !target || source === target) return
    nodes.add(source)
    nodes.add(target)
    links.push({ source, target, value: Math.abs(value) })
  })
  return { nodes: [...nodes].map((name) => ({ name })), links }
}

interface HierarchyNode { name: string; value?: number; children?: HierarchyNode[] }

function hierarchyData(data: ChartData): HierarchyNode[] {
  const roots: HierarchyNode[] = []
  data.categories.forEach((path, index) => {
    const parts = path.split('/').map((part) => part.trim()).filter(Boolean)
    if (parts.length === 0) return
    let siblings = roots
    for (const [partIndex, name] of parts.entries()) {
      let node = siblings.find((candidate) => candidate.name === name)
      if (!node) {
        node = { name }
        siblings.push(node)
      }
      if (partIndex === parts.length - 1) node.value = Math.abs(data.series[0]?.values[index] ?? 0)
      else {
        node.children ??= []
        siblings = node.children
      }
    }
  })
  return roots
}

function histogram(values: readonly (number | null)[]): { labels: string[]; counts: number[] } {
  const nums = values.filter((value): value is number => value !== null && Number.isFinite(value))
  if (nums.length === 0) return { labels: [], counts: [] }
  const min = Math.min(...nums)
  const max = Math.max(...nums)
  if (min === max) return { labels: [String(min)], counts: [nums.length] }
  const binCount = Math.max(1, Math.ceil(Math.log2(nums.length) + 1))
  const width = (max - min) / binCount
  const counts = Array.from({ length: binCount }, () => 0)
  nums.forEach((value) => {
    const index = Math.min(binCount - 1, Math.floor((value - min) / width))
    counts[index]++
  })
  const format = (value: number) => Number(value.toPrecision(6)).toString()
  const labels = counts.map((_, index) => `${format(min + index * width)}–${format(min + (index + 1) * width)}`)
  return { labels, counts }
}

function cumulativePercent(values: readonly (number | null)[]): number[] {
  const total = values.reduce<number>((sum, value) => sum + Math.max(0, value ?? 0), 0)
  let running = 0
  return values.map((value) => {
    running += Math.max(0, value ?? 0)
    return total === 0 ? 0 : running / total
  })
}

export function buildEChartsOption(
  spec: ChartSpec,
  data: ChartData,
  theme?: ChartTheme,
): Record<string, unknown> {
  const t = resolveTheme(theme)
  const fmt = valueFormatter(spec)
  const type = normalizeChartType(spec.type)
  const showLegend = spec.legend ?? data.series.length > 1
  const base: Record<string, unknown> = {
    color: t.palette,
    backgroundColor: t.background,
    animationDuration: 250,
    textStyle: { color: t.textColor },
    title: spec.title
      ? { text: spec.title, left: 'center', textStyle: { fontSize: 13, fontWeight: 600, color: t.textColor } }
      : undefined,
    legend: showLegend
      ? { bottom: 0, type: 'scroll', itemWidth: 12, itemHeight: 8, textStyle: { color: t.mutedTextColor } }
      : undefined,
    tooltip: {
      trigger: ['Pie', 'Donut', 'Funnel', 'Treemap', 'Sunburst', 'Gauge', 'Sankey', 'Relation', 'Chord', 'Heatmap', 'WordCloud'].includes(type) ? 'item' : 'axis',
      valueFormatter: fmt,
    },
  }
  const gridTop = spec.title ? 34 : 12
  const gridBottom = showLegend ? 28 : 8
  const grid = { left: 8, right: 16, top: gridTop, bottom: gridBottom, containLabel: true }

  switch (type) {
    case 'Pie':
    case 'Donut':
    case 'Funnel': {
      const slices = data.categories.map((name, i) => ({ name, value: data.series[0]?.values[i] ?? 0 }))
      if (type === 'Funnel') {
        return { ...base, series: [{ type: 'funnel', data: slices, top: gridTop + 8, bottom: gridBottom + 8, label: { color: t.textColor } }] }
      }
      return {
        ...base,
        series: [
          {
            type: 'pie',
            radius: type === 'Donut' ? ['45%', '72%'] : '72%',
            data: slices,
            label: { show: false },
          },
        ],
      }
    }

    case 'Treemap': {
      return {
        ...base,
        series: [
          {
            type: 'treemap',
            roam: false,
            breadcrumb: { show: false },
            data: data.categories.map((name, i) => ({ name, value: Math.abs(data.series[0]?.values[i] ?? 0) })),
            label: { color: '#fff' },
          },
        ],
      }
    }

    case 'Sunburst':
      return {
        ...base,
        series: [
          {
            type: 'sunburst',
            data: hierarchyData(data),
            radius: ['12%', '82%'],
            label: { rotate: 'radial' },
          },
        ],
      }

    case 'Gauge': {
      const index = data.series[0]?.values.findIndex((value) => value !== null) ?? -1
      const value = index >= 0 ? data.series[0].values[index] ?? 0 : 0
      return {
        ...base,
        legend: undefined,
        series: [
          {
            type: 'gauge',
            detail: { formatter: fmt },
            data: [{ name: data.categories[index] ?? data.series[0]?.name ?? 'Value', value }],
          },
        ],
      }
    }

    case 'WordCloud': {
      const values = data.series[0]?.values ?? []
      const magnitudes = values.filter((value): value is number => value !== null).map(Math.abs)
      const max = Math.max(1, ...magnitudes)
      const words = data.categories.flatMap((name, index) => {
        const value = values[index]
        return value === null || value === undefined ? [] : [[value, name, index] as [number, string, number]]
      })
      return {
        ...base,
        legend: undefined,
        tooltip: { trigger: 'item', formatter: (params: { value: [number, string] }) => `${params.value[1]}: ${fmt?.(params.value[0]) ?? params.value[0]}` },
        series: [
          {
            type: 'custom',
            coordinateSystem: 'none',
            data: words,
            renderItem: (_params: unknown, api: any) => {
              const value = Number(api.value(0))
              const name = String(api.value(1))
              const index = Number(api.value(2))
              const angle = index * 2.399963229728653
              const radius = Math.sqrt(index) * 18
              return {
                type: 'text',
                x: api.getWidth() / 2 + Math.cos(angle) * radius,
                y: api.getHeight() / 2 + Math.sin(angle) * radius,
                style: {
                  text: name,
                  fill: t.palette[index % t.palette.length],
                  fontSize: 11 + (Math.abs(value) / max) * 25,
                  align: 'center',
                  verticalAlign: 'middle',
                },
              }
            },
          },
        ],
      }
    }

    case 'Heatmap': {
      const cells: [number, number, number][] = []
      data.series.forEach((s, x) => {
        s.values.forEach((v, y) => {
          if (v !== null) cells.push([x, y, v])
        })
      })
      const flat = cells.map((c) => c[2])
      return {
        ...base,
        grid: { ...grid, right: 80 },
        xAxis: { type: 'category', data: data.series.map((s) => s.name), ...axisCommon(t) },
        yAxis: { type: 'category', data: data.categories, ...axisCommon(t) },
        visualMap: {
          min: flat.length ? Math.min(...flat) : 0,
          max: flat.length ? Math.max(...flat) : 1,
          orient: 'vertical',
          right: 4,
          top: 'center',
          textStyle: { color: t.mutedTextColor },
          inRange: { color: ['#e8f2ef', t.palette[1] ?? '#0fa98f'] },
        },
        series: [{ type: 'heatmap', data: cells, label: { show: false } }],
      }
    }

    case 'Radar': {
      const maxAll = Math.max(1, ...data.series.flatMap((s) => s.values.filter((v): v is number => v !== null)))
      return {
        ...base,
        radar: {
          indicator: data.categories.map((name) => ({ name, max: maxAll * 1.1 })),
          axisName: { color: t.mutedTextColor },
          splitLine: { lineStyle: { color: t.gridLineColor } },
          axisLine: { lineStyle: { color: t.gridLineColor } },
        },
        series: [
          {
            type: 'radar',
            data: data.series.map((s) => ({ name: s.name, value: s.values.map((v) => v ?? 0) })),
          },
        ],
      }
    }

    case 'Sankey': {
      // Sankey needs (source, target, value) triples. ChartData carries only
      // one text column (categories), so the convention is: categories =
      // "Source → Target" labels (a single text column with an arrow or ">"
      // separator) and series[0] = flow values. This keeps sankey usable from
      // the standard extraction until extract.ts learns multi-text-column
      // layouts (tracked in the roadmap alongside DrawingML work).
      const { nodes, links } = linkData(data)
      return {
        ...base,
        series: [
          {
            type: 'sankey',
            data: nodes,
            links,
            label: { color: t.textColor },
            lineStyle: { color: 'gradient', curveness: 0.5 },
          },
        ],
      }
    }

    case 'Relation':
    case 'Chord': {
      const { nodes, links } = linkData(data)
      return {
        ...base,
        series: [
          {
            type: 'graph',
            layout: type === 'Chord' ? 'circular' : 'force',
            circular: type === 'Chord' ? { rotateLabel: true } : undefined,
            force: type === 'Relation' ? { repulsion: 130, edgeLength: 80 } : undefined,
            roam: true,
            data: nodes,
            links,
            label: { show: true, color: t.textColor },
            lineStyle: { curveness: type === 'Chord' ? 0.25 : 0.08, opacity: 0.65 },
            edgeSymbol: type === 'Relation' ? ['none', 'arrow'] : ['none', 'none'],
          },
        ],
      }
    }

    case 'Boxplot': {
      const boxes = data.series
        .map((s) => ({ name: s.name, summary: fiveNumberSummary(s.values) }))
        .filter((b): b is { name: string; summary: [number, number, number, number, number] } => b.summary !== null)
      return {
        ...base,
        grid,
        xAxis: { type: 'category', data: boxes.map((b) => b.name), ...axisCommon(t) },
        yAxis: { type: 'value', ...axisCommon(t), axisLabel: { color: t.mutedTextColor, formatter: fmt } },
        series: [{ type: 'boxplot', data: boxes.map((b) => b.summary) }],
      }
    }

    case 'Candlestick': {
      const [open, close, low, high] = data.series
      const candles = data.categories.map((_, index) => [
        open?.values[index] ?? 0,
        close?.values[index] ?? 0,
        low?.values[index] ?? 0,
        high?.values[index] ?? 0,
      ])
      return {
        ...base,
        grid,
        xAxis: { type: 'category', data: data.categories, ...axisCommon(t) },
        yAxis: { type: 'value', ...axisCommon(t), axisLabel: { color: t.mutedTextColor, formatter: fmt } },
        series: [{ type: 'candlestick', name: 'OHLC', data: candles }],
      }
    }

    case 'Histogram': {
      const bins = histogram(data.series[0]?.values ?? [])
      return {
        ...base,
        legend: undefined,
        grid,
        xAxis: { type: 'category', data: bins.labels, ...axisCommon(t), axisLabel: { color: t.mutedTextColor, interval: 0, rotate: 30 } },
        yAxis: { type: 'value', minInterval: 1, ...axisCommon(t), axisLabel: { color: t.mutedTextColor } },
        series: [{ type: 'bar', name: 'Frequency', data: bins.counts, barWidth: '96%' }],
      }
    }

    case 'Pareto': {
      const values = data.series[0]?.values ?? []
      return {
        ...base,
        grid,
        xAxis: { type: 'category', data: data.categories, ...axisCommon(t) },
        yAxis: [
          { type: 'value', ...axisCommon(t), axisLabel: { color: t.mutedTextColor, formatter: fmt } },
          { type: 'value', min: 0, max: 1, position: 'right', ...axisCommon(t), splitLine: { show: false }, axisLabel: { color: t.mutedTextColor, formatter: (value: number) => `${Math.round(value * 100)}%` } },
        ],
        series: [
          { type: 'bar', name: data.series[0]?.name ?? 'Value', data: values, barMaxWidth: 26 },
          { type: 'line', name: 'Cumulative %', data: cumulativePercent(values), yAxisIndex: 1, smooth: false },
        ],
      }
    }

    case 'Waterfall': {
      const w = waterfallSegments(data.series[0]?.values ?? [])
      const categories = [...data.categories, 'Total']
      const riseColor = t.palette[1] ?? '#0fa98f'
      const fallColor = t.palette[3] ?? '#c0453b'
      const totalColor = t.palette[0] ?? '#2f7ed8'
      return {
        ...base,
        legend: undefined,
        grid,
        xAxis: { type: 'category', data: categories, ...axisCommon(t) },
        yAxis: { type: 'value', ...axisCommon(t), axisLabel: { color: t.mutedTextColor, formatter: fmt } },
        series: [
          // Invisible base stack positions each floating segment.
          {
            type: 'bar',
            stack: 'wf',
            itemStyle: { color: 'transparent' },
            emphasis: { itemStyle: { color: 'transparent' } },
            tooltip: { show: false },
            data: [...w.base, 0],
          },
          {
            type: 'bar',
            stack: 'wf',
            name: 'Increase',
            itemStyle: { color: riseColor },
            data: [...w.rise, null],
          },
          {
            type: 'bar',
            stack: 'wf',
            name: 'Decrease',
            itemStyle: { color: fallColor },
            data: [...w.fall, null],
          },
          {
            type: 'bar',
            stack: 'wf',
            name: 'Total',
            itemStyle: { color: totalColor },
            data: [...w.rise.map(() => null), w.total],
          },
        ],
      }
    }

    case 'Scatter':
      return {
        ...base,
        grid: { left: 44, right: 16, top: gridTop, bottom: gridBottom + 20, containLabel: false },
        xAxis: { type: 'value', ...axisCommon(t) },
        yAxis: { type: 'value', ...axisCommon(t), axisLabel: { color: t.mutedTextColor, formatter: fmt } },
        series:
          data.series.length >= 2
            ? [
                {
                  type: 'scatter',
                  name: data.series[1].name,
                  symbolSize: 8,
                  data: data.series[0].values.map((x, i) => [x, data.series[1].values[i]]),
                },
              ]
            : [
                {
                  type: 'scatter',
                  name: data.series[0]?.name ?? 'Series 1',
                  symbolSize: 8,
                  data: (data.series[0]?.values ?? []).map((y, i) => [i, y]),
                },
              ],
      }

    case 'Bubble': {
      const [x, y, size] = data.series
      const points = data.categories.map((category, index) => [
        x?.values[index] ?? null,
        y?.values[index] ?? null,
        Math.abs(size?.values[index] ?? 0),
        category,
      ])
      const maxSize = Math.max(1, ...points.map((point) => Number(point[2])))
      return {
        ...base,
        grid: { left: 44, right: 16, top: gridTop, bottom: gridBottom + 20, containLabel: false },
        xAxis: { type: 'value', name: x?.name, ...axisCommon(t) },
        yAxis: { type: 'value', name: y?.name, ...axisCommon(t), axisLabel: { color: t.mutedTextColor, formatter: fmt } },
        series: [
          {
            type: 'scatter',
            name: size?.name ?? 'Size',
            data: points,
            symbolSize: (point: [number, number, number]) => 8 + 34 * Math.sqrt(point[2] / maxSize),
          },
        ],
      }
    }

    case 'ColumnStacked':
    case 'ColumnPercentStacked':
    case 'BarStacked':
    case 'BarPercentStacked':
    case 'AreaStacked':
    case 'AreaPercentStacked': {
      const horizontal = type === 'BarStacked' || type === 'BarPercentStacked'
      const area = type === 'AreaStacked' || type === 'AreaPercentStacked'
      const percent = type === 'ColumnPercentStacked' || type === 'BarPercentStacked' || type === 'AreaPercentStacked'
      const plotted = percent ? percentSeries(data) : data.series
      const categoryAxis = { type: 'category', data: data.categories, ...axisCommon(t) }
      const valueAxis = {
        type: 'value',
        ...(percent ? { min: -1, max: 1 } : {}),
        ...axisCommon(t),
        axisLabel: {
          color: t.mutedTextColor,
          formatter: percent ? (value: number) => `${Math.round(value * 100)}%` : fmt,
        },
      }
      return {
        ...base,
        grid,
        xAxis: horizontal ? valueAxis : categoryAxis,
        yAxis: horizontal ? categoryAxis : valueAxis,
        series: plotted.map((series) => ({
          type: area ? 'line' : 'bar',
          name: series.name,
          data: series.values,
          stack: 'total',
          barMaxWidth: 26,
          ...(area ? { areaStyle: { opacity: 0.32 }, showSymbol: false } : {}),
        })),
      }
    }

    case 'Bar': // horizontal bars — no combo/secondary/trendline (axis flip)
      return {
        ...base,
        grid,
        xAxis: { type: 'value', ...axisCommon(t), axisLabel: { color: t.mutedTextColor, formatter: fmt } },
        yAxis: { type: 'category', data: data.categories, ...axisCommon(t) },
        series: data.series.map((s) => ({ type: 'bar', name: s.name, data: s.values, barMaxWidth: 26 })),
      }

    case 'Column':
    case 'Line':
    case 'Area':
    case 'Combination':
    default: {
      // The categorical family: per-series combo types, secondary axis,
      // trendlines. This is the workhorse path.
      const baseMark = type === 'Column' || type === 'Combination' ? 'bar' : 'line'
      const overrides = spec.series ?? {}
      const anySecondary = data.series.some((s) => overrides[s.name]?.secondaryAxis)

      const yAxes: Record<string, unknown>[] = [
        { type: 'value', ...axisCommon(t), axisLabel: { color: t.mutedTextColor, formatter: fmt } },
      ]
      if (anySecondary) {
        yAxes.push({
          type: 'value',
          position: 'right',
          ...axisCommon(t),
          splitLine: { show: false },
          axisLabel: { color: t.mutedTextColor, formatter: fmt },
        })
      }

      const series: Record<string, unknown>[] = []
      data.series.forEach((s, idx) => {
        const o = overrides[s.name] ?? {}
        const mark = o.type
          ? (o.type === 'column' ? 'bar' : 'line')
          : type === 'Combination' && idx % 2 === 1
            ? 'line'
            : baseMark
        const isArea = o.type === 'area' || (!o.type && type === 'Area')
        series.push({
          type: mark,
          name: s.name,
          data: s.values,
          yAxisIndex: o.secondaryAxis && anySecondary ? 1 : 0,
          barMaxWidth: 26,
          smooth: false,
          ...(isArea ? { areaStyle: { opacity: 0.25 } } : {}),
        })
        if (o.trendline) {
          const trend =
            o.trendline === 'movingAverage'
              ? movingAverage(s.values, o.trendlineWindow ?? 3)
              : linearTrend(s.values)
          series.push({
            type: 'line',
            name: `${s.name} (${o.trendline === 'movingAverage' ? 'MA' : 'trend'})`,
            data: trend,
            yAxisIndex: o.secondaryAxis && anySecondary ? 1 : 0,
            showSymbol: false,
            lineStyle: { type: 'dashed', width: 1.5, color: t.palette[idx % t.palette.length] },
            itemStyle: { color: t.palette[idx % t.palette.length] },
            tooltip: { show: false },
            silent: true,
          })
        }
      })

      return {
        ...base,
        grid: anySecondary ? { ...grid, right: 8 } : grid,
        xAxis: { type: 'category', data: data.categories, axisTick: { alignWithLabel: true }, ...axisCommon(t) },
        yAxis: yAxes,
        series,
      }
    }
  }
}
