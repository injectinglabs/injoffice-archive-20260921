import { useEffect, useRef, useState } from 'react'
import { LocaleType } from '@univerjs/core'
import {
  ChartManager,
  ChartFloat,
  ChartPanel,
  CHART_COMPONENT_KEY,
  CHART_TYPES,
  type ChartType,
} from '@injoffice/charts'
import { PivotManager, PivotPanel } from '@injoffice/pivots'
import { ShapeManager, ShapeFloat, SHAPE_COMPONENT_KEY, SHAPE_KINDS, type ShapeKind } from '@injoffice/shapes'
import { ConnectorManager, DataPanel, jsonToGrid, csvToGrid, type ConnectorSource } from '@injoffice/connectors'
import { createOssUniver, createSheetsPresetBundle, type SheetsFeatureConfig } from '@injoffice/univer-sheets/browser'
import { bindUniverColorScheme, univerDarkMode } from './univerColorScheme'

async function playgroundFetchSource(source: ConnectorSource): Promise<unknown[][]> {
  const res = await fetch(source.url)
  if (!res.ok) throw new Error(`fetch failed (${res.status})`)
  if (source.format === 'csv') return csvToGrid(await res.text())
  return jsonToGrid(await res.json(), source.path)
}

const SAMPLE = [
  ['Quarter', 'Revenue', 'Costs', 'Profit'],
  ['Q1', 120, 80, 40],
  ['Q2', 150, 90, 60],
  ['Q3', 170, 95, 75],
  ['Q4', 210, 110, 100],
]

function buildCellData() {
  const cellData: Record<number, Record<number, { v: string | number }>> = {}
  SAMPLE.forEach((row, r) => {
    cellData[r] = {}
    row.forEach((v, c) => {
      cellData[r][c] = { v }
    })
  })
  return cellData
}

/** Univer OSS editor shell: charts/pivots/shapes. Not native XLSX paint. */
export function UniverEditorShell({ features }: { features?: SheetsFeatureConfig } = {}) {
  const containerRef = useRef<HTMLDivElement | null>(null)
  const managerRef = useRef<ChartManager | null>(null)
  const pivotsRef = useRef<PivotManager | null>(null)
  const shapesRef = useRef<ShapeManager | null>(null)
  const connectorsRef = useRef<ConnectorManager | null>(null)
  const [shapeKind, setShapeKind] = useState<ShapeKind>('rect')
  const [type, setType] = useState<ChartType>('column')
  const [ready, setReady] = useState(false)

  useEffect(() => {
    const formulaWorker = new Worker(new URL('./formula-worker.ts', import.meta.url), { type: 'module' })
    const bundle = createSheetsPresetBundle({
      container: containerRef.current!,
      workerURL: formulaWorker,
      features,
    })
    const { univer, univerAPI } = createOssUniver({
      locale: LocaleType.EN_US,
      locales: { [LocaleType.EN_US]: bundle.locale },
      presets: bundle.presets,
      darkMode: univerDarkMode(),
    })
    const unbindScheme = bindUniverColorScheme(univerAPI)
    univerAPI.createWorkbook({ id: 'wb1', sheets: { s1: { id: 's1', name: 'Data', cellData: buildCellData() } } })
    univerAPI.registerComponent(CHART_COMPONENT_KEY, ChartFloat)
    univerAPI.registerComponent(SHAPE_COMPONENT_KEY, ShapeFloat)

    const manager = new ChartManager(univerAPI)
    manager.start()
    managerRef.current = manager
    const pivots = new PivotManager(univerAPI)
    pivots.start()
    pivotsRef.current = pivots
    const shapes = new ShapeManager(univerAPI)
    shapes.start()
    shapesRef.current = shapes
    const connectors = new ConnectorManager(univerAPI, playgroundFetchSource)
    connectorsRef.current = connectors
    ;(window as unknown as Record<string, unknown>).__injoffice = { univerAPI, charts: manager, pivots, shapes, connectors }
    setReady(true)

    return () => {
      unbindScheme()
      manager.stop()
      pivots.stop()
      shapes.stop()
      univer.dispose()
      formulaWorker.terminate()
    }
  }, [features])

  return (
    <div className="platen-fill" data-demo-surface="univer-sandbox">
      <div className="toolstrip univer-toolbar" role="toolbar" aria-label="Workbook tools">
        <div className="univer-toolbar__group" role="group" aria-label="Chart tools">
          <label className="tool-field">
            <span>Chart</span>
            <select value={type} onChange={(e) => setType(e.target.value as ChartType)} aria-label="Chart type">
              {CHART_TYPES.map((t) => (
                <option key={t} value={t}>{t}</option>
              ))}
            </select>
          </label>
          <button
            type="button"
            className="workbench-button workbench-button--primary"
            disabled={!ready}
            onClick={() => {
              const spec = managerRef.current?.createFromSelection(type)
              if (!spec) window.alert('Select a data range first')
            }}
          >
            Add chart
          </button>
          <button
            type="button"
            className="workbench-button workbench-button--danger"
            disabled={!ready}
            onClick={() => {
              const charts = managerRef.current?.list() ?? []
              const last = charts[charts.length - 1]
              if (last) managerRef.current?.remove(last.id)
            }}
          >
            Remove last chart
          </button>
        </div>
        <div className="univer-toolbar__group" role="group" aria-label="Pivot table tools">
          <button
            type="button"
            className="workbench-button"
            disabled={!ready}
            onClick={() => {
              const spec = pivotsRef.current?.createFromSelection()
              if (!spec) window.alert('Select a data block including its header row first')
            }}
          >
            Add pivot table
          </button>
        </div>
        <div className="univer-toolbar__group" role="group" aria-label="Shape tools">
          <label className="tool-field">
            <span>Shape</span>
            <select value={shapeKind} onChange={(e) => setShapeKind(e.target.value as ShapeKind)} aria-label="Shape kind">
              {SHAPE_KINDS.map((k) => (
                <option key={k} value={k}>{k}</option>
              ))}
            </select>
          </label>
          <button
            type="button"
            className="workbench-button"
            disabled={!ready}
            onClick={() => shapesRef.current?.create(shapeKind)}
          >
            Add shape
          </button>
        </div>
        <span className="univer-toolbar__status">
          Select A1:D5 to add a chart · charts update when cells change
        </span>
      </div>
      <div className="split">
        <div ref={containerRef} className="split-main" />
        {ready && managerRef.current && pivotsRef.current && (
          <aside className="split-side">
            <ChartPanel manager={managerRef.current} />
            <PivotPanel manager={pivotsRef.current} />
            {connectorsRef.current && <DataPanel manager={connectorsRef.current} />}
          </aside>
        )}
      </div>
    </div>
  )
}
