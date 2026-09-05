import { useEffect, useRef, useState } from 'react'
import { LocaleType } from '@univerjs/core'
import {
  ChartManager,
  ChartFloat,
  ChartPanel,
  CHART_COMPONENT_KEY,
  CHART_TYPES,
  specsFromFileCharts,
  type ChartType,
  type FileChartAnchor,
  type FileChartInfo,
} from '@injoffice/charts'
import top100 from './top100Workbook.json'
import { PivotManager, PivotPanel } from '@injoffice/pivots'
import { ShapeManager, ShapeFloat, SHAPE_COMPONENT_KEY, SHAPE_KINDS, type ShapeKind } from '@injoffice/shapes'
import { ConnectorManager, DataPanel, jsonToGrid, csvToGrid, type ConnectorSource } from '@injoffice/connectors'
import type { SparklineCommandController } from '../../../packages/sparklines/src/commands'
import type { OutlineCommandController } from '../../../packages/outlines/src/commands'
import { SheetsPowerPanels } from './sheetsPowerPanels'
import {
  createSheetsOutlineController,
  createSheetsSparklineController,
  outlineGroupFromRange,
  sparklineInputFromRange,
} from './sheetsPower'
import {
  createOssUniver,
  createSheetsPresetBundle,
  registerInjOfficeInsertMenu,
  type InjOfficeInsertFeatureConfig,
  type SheetsFeatureConfig,
} from '@injoffice/univer-sheets/browser'
import { bindUniverColorScheme, univerDarkMode } from './univerColorScheme'
import { deferNestedReactRootStart } from './nestedReactRootLifecycle'
import '@injoffice/univer-sheets/styles.css'

async function playgroundFetchSource(source: ConnectorSource): Promise<unknown[][]> {
  const res = await fetch(source.url)
  if (!res.ok) throw new Error(`fetch failed (${res.status})`)
  if (source.format === 'csv') return csvToGrid(await res.text())
  return jsonToGrid(await res.json(), source.path)
}

function workbookCellData() {
  const cellData: Record<number, Record<number, { v?: string | number; f?: string }>> = {}
  for (const [row, cols] of Object.entries(top100.cellData)) {
    const next: Record<number, { v?: string | number; f?: string }> = {}
    for (const [col, cell] of Object.entries(cols)) {
      next[Number(col)] = cell as { v?: string | number; f?: string }
    }
    cellData[Number(row)] = next
  }
  return cellData
}

export interface UniverEditorProps {
  /** Optional Univer OSS capabilities. Every feature is enabled by default. */
  features?: SheetsFeatureConfig
  /** Optional InjOffice Insert commands. Chart, pivot table, and shape default to enabled. */
  insertFeatures?: InjOfficeInsertFeatureConfig
}

export default function UniverEditor({ features, insertFeatures }: UniverEditorProps) {
  const containerRef = useRef<HTMLDivElement | null>(null)
  const managerRef = useRef<ChartManager | null>(null)
  const pivotsRef = useRef<PivotManager | null>(null)
  const shapesRef = useRef<ShapeManager | null>(null)
  const connectorsRef = useRef<ConnectorManager | null>(null)
  const sparklineRef = useRef<SparklineCommandController | null>(null)
  const outlineRef = useRef<OutlineCommandController | null>(null)
  const apiRef = useRef<{
    getActiveWorkbook(): {
      getActiveSheet(): { getSheetId(): string; getActiveRange(): { getRange(): { startRow: number; startColumn: number; endRow: number; endColumn: number } } | null } | null
      getSheetBySheetId(id: string): {
        hideRows(start: number, count: number): void
        showRows(start: number, count: number): void
        hideColumns(start: number, count: number): void
        showColumns(start: number, count: number): void
      } | null
    } | null
  } | null>(null)
  const [shapeKind, setShapeKind] = useState<ShapeKind>('rect')
  const [type, setType] = useState<ChartType>('column')
  const [ready, setReady] = useState(false)
  const chartTypeRef = useRef(type)
  const shapeKindRef = useRef(shapeKind)
  chartTypeRef.current = type
  shapeKindRef.current = shapeKind

  useEffect(() => deferNestedReactRootStart(() => {
    if (!containerRef.current) return
    const bundle = createSheetsPresetBundle({
      container: containerRef.current,
      features,
    })
    const { univer, univerAPI } = createOssUniver({
      locale: LocaleType.EN_US,
      locales: { [LocaleType.EN_US]: bundle.locale },
      presets: bundle.presets,
      darkMode: univerDarkMode(),
    })
    const unbindScheme = bindUniverColorScheme(univerAPI)
    univerAPI.createWorkbook({
      id: 'wb1',
      name: 'top_100_companies.xlsx',
      sheets: {
        [top100.sheetId]: {
          id: top100.sheetId,
          name: top100.sheetName,
          cellData: workbookCellData(),
          columnData: {
            0: { w: 72 },
            1: { w: 220 },
            2: { w: 88 },
            3: { w: 160 },
            4: { w: 140 },
            5: { w: 140 },
            6: { w: 120 },
          },
        },
      },
    })
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
    apiRef.current = univerAPI as typeof apiRef.current
    const sparkline = createSheetsSparklineController()
    sparklineRef.current = sparkline
    const outline = createSheetsOutlineController({
      hide(sheetId, axis, start, count) {
        const sheet = univerAPI.getActiveWorkbook()?.getSheetBySheetId(sheetId)
        if (!sheet) return
        if (axis === 'row') sheet.hideRows(start, count)
        else sheet.hideColumns(start, count)
      },
      show(sheetId, axis, start, count) {
        const sheet = univerAPI.getActiveWorkbook()?.getSheetBySheetId(sheetId)
        if (!sheet) return
        if (axis === 'row') sheet.showRows(start, count)
        else sheet.showColumns(start, count)
      },
    })
    outlineRef.current = outline
    const insertMenu = registerInjOfficeInsertMenu(
      univer,
      {
        chart: () => {
          if (manager.createFromSelection(chartTypeRef.current)) return true
          window.alert('Select a data range first')
          return false
        },
        pivotTable: () => {
          if (pivots.createFromSelection()) return true
          window.alert('Select a data block including its header row first')
          return false
        },
        shape: () => shapes.create(shapeKindRef.current) !== null,
      },
      insertFeatures,
    )
    ;(window as unknown as Record<string, unknown>).__injoffice = { univerAPI, charts: manager, pivots, shapes, connectors, sparklines: sparkline, outlines: outline }
    setReady(true)

    let cancelled = false
    let frames = 0
    const mountFileCharts = () => {
      if (cancelled) return
      frames++
      const canvasMounted = containerRef.current?.querySelector('canvas')?.isConnected === true
      if (!canvasMounted) {
        if (frames < 60) requestAnimationFrame(mountFileCharts)
        else console.warn('InjOffice playground: skipped file charts because the Univer render engine did not mount')
        return
      }
      const sheets = univerAPI.getActiveWorkbook()?.getSheets() ?? []
      if (sheets.length === 0 && frames < 60) {
        requestAnimationFrame(mountFileCharts)
        return
      }
      const sheetIdByName: Record<string, string> = {}
      for (const sh of sheets) sheetIdByName[sh.getSheetName()] = sh.getSheetId()
      const { conversions, skipped } = specsFromFileCharts(
        top100.charts as FileChartInfo[],
        top100.anchors as Record<string, FileChartAnchor>,
        sheetIdByName,
      )
      for (const conversion of conversions) manager.add(conversion.spec, conversion.cellAnchor)
      if (manager.list().length === 0 && conversions.length > 0 && frames < 60) {
        requestAnimationFrame(mountFileCharts)
        return
      }
      if (skipped) console.warn(`InjOffice playground: skipped ${skipped} unrenderable file chart(s)`)
      if (manager.list().length === 0) {
        console.warn('InjOffice playground: file chart did not mount after the sheet skeleton wait', {
          sheetIdByName,
          conversions,
        })
      }
    }
    // Float DOMs subscribe to the sheet render engine's client rect. Workbook
    // model creation happens before Univer mounts that engine, so hydrate file
    // charts only after Univer's UI lifecycle reaches Rendered.
    const renderedHook = univerAPI.getHooks().onRendered(() => {
      requestAnimationFrame(mountFileCharts)
    })

    return () => {
      cancelled = true
      renderedHook.dispose()
      unbindScheme()
      manager.stop()
      pivots.stop()
      shapes.stop()
      insertMenu.dispose()
      // Univer owns a nested React root. Dispose it after React finishes the
      // outer route commit so both roots are never unmounted synchronously.
      window.setTimeout(() => univer.dispose(), 0)
    }
  }), [features, insertFeatures])

  return (
    <div className="platen-fill" data-demo-surface="sheets">
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
        <div className="univer-toolbar__group" role="group" aria-label="Sparkline and outline tools">
          <button
            type="button"
            className="workbench-button"
            disabled={!ready}
            onClick={() => {
              const sheet = apiRef.current?.getActiveWorkbook?.()?.getActiveSheet?.()
              const selected = sheet?.getActiveRange?.()?.getRange()
              if (!sheet || !selected) {
                window.alert('Select a one-row or one-column range first')
                return
              }
              const input = sparklineInputFromRange('line', sheet.getSheetId(), selected)
              if (!input) {
                window.alert('Sparklines need a single row or column')
                return
              }
              try { sparklineRef.current?.create(input) } catch (reason) {
                window.alert(reason instanceof Error ? reason.message : String(reason))
              }
            }}
          >
            Add sparkline
          </button>
          <button
            type="button"
            className="workbench-button"
            disabled={!ready}
            onClick={() => {
              const sheet = apiRef.current?.getActiveWorkbook?.()?.getActiveSheet?.()
              const selected = sheet?.getActiveRange?.()?.getRange()
              if (!sheet || !selected) {
                window.alert('Select rows to group first')
                return
              }
              const result = outlineRef.current?.add(outlineGroupFromRange(`rows-${Date.now().toString(36)}`, sheet.getSheetId(), 'row', selected))
              if (!result?.ok) window.alert(result?.issues.map((issue) => issue.message).join('; ') || 'Could not group rows')
            }}
          >
            Group rows
          </button>
          <button
            type="button"
            className="workbench-button"
            disabled={!ready}
            onClick={() => {
              const sheet = apiRef.current?.getActiveWorkbook?.()?.getActiveSheet?.()
              const selected = sheet?.getActiveRange?.()?.getRange()
              if (!sheet || !selected) {
                window.alert('Select columns to group first')
                return
              }
              const result = outlineRef.current?.add(outlineGroupFromRange(`cols-${Date.now().toString(36)}`, sheet.getSheetId(), 'column', selected))
              if (!result?.ok) window.alert(result?.issues.map((issue) => issue.message).join('; ') || 'Could not group columns')
            }}
          >
            Group columns
          </button>
        </div>
        <span className="univer-toolbar__status">Edit in Univer · file round trips use the local Go sidecar</span>
      </div>
      <div className="split">
        <div ref={containerRef} className="split-main" />
        {ready && managerRef.current && pivotsRef.current && sparklineRef.current && outlineRef.current && shapesRef.current && (
          <aside className="split-side">
            <ChartPanel manager={managerRef.current} />
            <PivotPanel manager={pivotsRef.current} />
            {connectorsRef.current && <DataPanel manager={connectorsRef.current} />}
            <SheetsPowerPanels
              sparkline={sparklineRef.current}
              outline={outlineRef.current}
              shapes={shapesRef.current}
              sheetId={top100.sheetId}
            />
          </aside>
        )}
      </div>
    </div>
  )
}
