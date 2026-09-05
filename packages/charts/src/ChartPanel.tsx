import { useEffect, useMemo, useReducer, useState } from 'react'
import { ChartCommandController } from './commands'
import { ChartImageExportManager } from './imageExport'
import type { ChartImageExportArtifact, ChartImageFormat } from './imageExport'
import type { ChartManager } from './manager'
import { CHART_TYPES, normalizeChartType } from './types'
import type { ChartLayerOperation, ChartSpec, ChartType, SeriesComboType, TrendlineKind, ValueFormat } from './types'

// ChartPanel — the chart configuration side panel. Mutations flow through the
// public command facade so a host-supplied controller records snapshot undo.

const CATEGORICAL = new Set(['Line', 'Column', 'Area', 'Combination'])
const COMBO_TYPES: (SeriesComboType | '')[] = ['', 'column', 'line', 'area']
const TRENDLINES: (TrendlineKind | '')[] = ['', 'linear', 'movingAverage']
const FORMATS: ValueFormat[] = ['auto', 'plain', 'percent', 'currency']
const IMAGE_FORMATS: ChartImageFormat[] = ['png', 'jpeg', 'svg']
const LAYERS: Array<{ operation: ChartLayerOperation; label: string }> = [
  { operation: 'bringToFront', label: 'Bring to front' },
  { operation: 'bringForward', label: 'Bring forward' },
  { operation: 'sendBackward', label: 'Send backward' },
  { operation: 'sendToBack', label: 'Send to back' },
]

function TriState({
  label,
  value,
  onChange,
}: {
  label: string
  value: boolean | undefined
  onChange: (v: boolean | undefined) => void
}) {
  return (
    <label className="ioc-field">
      <span>{label}</span>
      <select
        value={value === undefined ? 'auto' : value ? 'yes' : 'no'}
        onChange={(e) => onChange(e.target.value === 'auto' ? undefined : e.target.value === 'yes')}
      >
        <option value="auto">Auto</option>
        <option value="yes">Yes</option>
        <option value="no">No</option>
      </select>
    </label>
  )
}

export interface ChartPanelCommonProps {
  exporter?: ChartImageExportManager
  /** Receives the validated artifact. When omitted, the browser downloads it. */
  onExport?: (artifact: ChartImageExportArtifact) => void | Promise<void>
  showLayerControls?: boolean
  showExportControls?: boolean
}

export type ChartPanelProps = ChartPanelCommonProps & (
  | { controller: ChartCommandController; manager?: never }
  | { /** @deprecated Pass the controller returned by registerUniverChartCommands to enable host undo. */ manager: ChartManager; controller?: never }
)

export function createChartPanelCommandBindings(controller: ChartCommandController, id: string) {
  return {
    remove: () => controller.remove(id),
    update: (patch: Partial<Omit<ChartSpec, 'id' | 'nativeIdentity'>>) => controller.update(id, patch),
    layer: (operation: ChartLayerOperation) => controller.layer(id, operation),
  }
}

export async function exportChartFromPanel(
  exporter: ChartImageExportManager,
  chartId: string,
  format: ChartImageFormat,
  deliver: (artifact: ChartImageExportArtifact) => void | Promise<void> = downloadChartImageArtifact,
): Promise<ChartImageExportArtifact> {
  const artifact = await exporter.export(chartId, { format })
  await deliver(artifact)
  return artifact
}

export function downloadChartImageArtifact(artifact: ChartImageExportArtifact): void {
  if (typeof document === 'undefined' || typeof URL?.createObjectURL !== 'function') throw new Error('Chart image download requires a browser document')
  const extension = artifact.format === 'jpeg' ? 'jpg' : artifact.format
  const safeId = artifact.chartId.replace(/[^A-Za-z0-9._-]+/g, '-').replace(/^[.-]+|[.-]+$/g, '') || 'chart'
  const url = URL.createObjectURL(new Blob([new Uint8Array(artifact.bytes)], { type: artifact.mediaType }))
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = `${safeId}.${extension}`
  try { anchor.click() } finally { URL.revokeObjectURL(url) }
}

export function ChartPanel({
  controller: suppliedController,
  manager: legacyManager,
  exporter,
  onExport,
  showLayerControls = true,
  showExportControls = true,
}: ChartPanelProps) {
  const controller = useMemo(() => {
    if (suppliedController) return suppliedController
    if (legacyManager) return new ChartCommandController(legacyManager)
    throw new TypeError('ChartPanel requires a controller or manager')
  }, [suppliedController, legacyManager])
  const manager = controller.manager
  const [, force] = useReducer((x: number) => x + 1, 0)
  const [selectedId, setSelectedId] = useState<string | null>(null)
  const [imageFormat, setImageFormat] = useState<ChartImageFormat>('png')
  const [exportState, setExportState] = useState<'idle' | 'working' | 'done' | 'error'>('idle')
  const [exportError, setExportError] = useState('')

  useEffect(() => manager.onChange(force), [manager])

  const charts = manager.list()
  const active: ChartSpec | undefined =
    (selectedId ? manager.getSpec(selectedId) : undefined) ?? charts[charts.length - 1]
  const seriesNames = useMemo(
    () => (active ? manager.getSeriesNames(active.id) : []),
    // charts.length re-derives after add/remove; force() covers spec edits.
    [manager, active?.id, charts.length],
  )

  if (charts.length === 0) {
    return <div className="ioc-panel ioc-panel--empty">No charts yet — select a range and insert one.</div>
  }
  if (!active) return null

  const commands = createChartPanelCommandBindings(controller, active.id)
  const patch = commands.update
  const patchSeries = (name: string, p: Partial<NonNullable<ChartSpec['series']>[string]>) => {
    const next = { ...(active.series ?? {}) }
    next[name] = { ...next[name], ...p }
    patch({ series: next })
  }

  return (
    <div className="ioc-panel">
      <div className="ioc-row">
        <select
          className="ioc-chart-select"
          value={active.id}
          onChange={(e) => setSelectedId(e.target.value)}
          aria-label="Chart"
        >
          {charts.map((c, i) => (
            <option key={c.id} value={c.id}>
              {c.title || `Chart ${i + 1} (${c.type})`}
            </option>
          ))}
        </select>
        <button type="button" className="ioc-remove" onClick={commands.remove}>
          Remove
        </button>
      </div>

      <label className="ioc-field">
        <span>Type</span>
        <select value={normalizeChartType(active.type)} onChange={(e) => patch({ type: e.target.value as ChartType })}>
          {CHART_TYPES.map((t) => (
            <option key={t} value={t}>{t}</option>
          ))}
        </select>
      </label>

      <label className="ioc-field">
        <span>Title</span>
        <input
          value={active.title ?? ''}
          placeholder="No title"
          onChange={(e) => patch({ title: e.target.value || undefined })}
        />
      </label>

      <TriState label="Legend" value={active.legend} onChange={(legend) => patch({ legend })} />
      <TriState
        label="First row is header"
        value={active.firstRowIsHeader}
        onChange={(firstRowIsHeader) => patch({ firstRowIsHeader })}
      />
      <TriState
        label="First column is category"
        value={active.firstColumnIsCategory}
        onChange={(firstColumnIsCategory) => patch({ firstColumnIsCategory })}
      />

      <label className="ioc-field">
        <span>Value format</span>
        <select
          value={active.valueFormat ?? 'auto'}
          onChange={(e) => patch({ valueFormat: e.target.value as ValueFormat })}
        >
          {FORMATS.map((f) => (
            <option key={f} value={f}>{f}</option>
          ))}
        </select>
      </label>

      {showLayerControls && <div className="ioc-series">
        <div className="ioc-series-head">Layer</div>
        <div className="ioc-chart-layer-actions" role="group" aria-label="Chart layer order">
          {LAYERS.map(({ operation, label }) => (
            <button
              key={operation}
              type="button"
              disabled={!manager.canLayer(active.id, operation)}
              onClick={() => commands.layer(operation)}
            >
              {label}
            </button>
          ))}
        </div>
      </div>}

      {exporter && showExportControls && (
        <div className="ioc-series">
          <div className="ioc-series-head">Export image</div>
          <div className="ioc-row">
            <select aria-label="Export format" value={imageFormat} onChange={(event) => setImageFormat(event.target.value as ChartImageFormat)}>
              {IMAGE_FORMATS.map((format) => <option key={format} value={format}>{format.toUpperCase()}</option>)}
            </select>
            <button
              type="button"
              disabled={exportState === 'working'}
              onClick={() => {
                setExportState('working')
                setExportError('')
                void exportChartFromPanel(exporter, active.id, imageFormat, onExport ?? downloadChartImageArtifact)
                  .then(() => setExportState('done'), (cause) => {
                    setExportState('error')
                    setExportError(cause instanceof Error ? cause.message : 'The renderer did not return an image.')
                  })
              }}
            >
              {exportState === 'working' ? 'Exporting…' : 'Export'}
            </button>
          </div>
          <div aria-live="polite" role="status">
            {exportState === 'done' ? 'Chart image exported.' : exportState === 'error' ? `Export failed: ${exportError}` : ''}
          </div>
        </div>
      )}

      {CATEGORICAL.has(normalizeChartType(active.type)) && seriesNames.length > 0 && (
        <div className="ioc-series">
          <div className="ioc-series-head">Series</div>
          {seriesNames.map((name) => {
            const o = active.series?.[name] ?? {}
            return (
              <div key={name} className="ioc-series-row">
                <span className="ioc-series-name" title={name}>{name}</span>
                <select
                  aria-label={`${name} type`}
                  title="Series type (combo)"
                  value={o.type ?? ''}
                  onChange={(e) =>
                    patchSeries(name, { type: (e.target.value || undefined) as SeriesComboType | undefined })
                  }
                >
                  {COMBO_TYPES.map((t) => (
                    <option key={t} value={t}>{t === '' ? 'default' : t}</option>
                  ))}
                </select>
                <label title="Secondary (right) axis">
                  <input
                    type="checkbox"
                    checked={o.secondaryAxis ?? false}
                    onChange={(e) => patchSeries(name, { secondaryAxis: e.target.checked || undefined })}
                  />
                  2nd axis
                </label>
                <select
                  aria-label={`${name} trendline`}
                  title="Trendline"
                  value={o.trendline ?? ''}
                  onChange={(e) =>
                    patchSeries(name, { trendline: (e.target.value || undefined) as TrendlineKind | undefined })
                  }
                >
                  {TRENDLINES.map((t) => (
                    <option key={t} value={t}>
                      {t === '' ? 'no trend' : t === 'movingAverage' ? 'moving avg' : 'linear trend'}
                    </option>
                  ))}
                </select>
              </div>
            )
          })}
        </div>
      )}
    </div>
  )
}
