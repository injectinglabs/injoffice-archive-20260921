import type { Nullable } from '@univerjs/core'
import type { FUniver } from '@univerjs/core/lib/facade'
// Side-effect type imports: FUniver's sheet/drawing/ui surface (getActiveWorkbook,
// addFloatDomToPosition, registerComponent, …) exists only as module
// augmentations contributed by these packages — without importing them the
// bare FUniver from core typechecks as method-less.
import type {} from '@univerjs/sheets/lib/facade'
import type {} from '@univerjs/sheets-ui/lib/facade'
import type {} from '@univerjs/sheets-drawing-ui/lib/facade'
import type {} from '@univerjs/ui/lib/facade'
import { extractChartData } from './extract'
import { invalidateMode, sourceNeedsRefresh } from './invalidate'
import { buildEChartsOption } from './option'
import { dropChart, publishChartOption } from './registry'
import type { ChartLayerOperation, ChartSpec, ChartTheme, ChartType, CellRangeRef } from './types'
import type { FileChartAnchor } from './fromFile'

// ChartManager — the host-facing API of @injoffice/charts.
//
// Owns the spec list, keeps every chart's rendered option in sync with the
// live cell data (one debounced refresh per Univer command batch, and only
// for charts whose source range intersects the edited cells), and mounts
// each chart as a Univer float-dom (movable/resizable via allowTransform).
//
// Deliberately storage-agnostic: serialize()/hydrate() move plain spec lists
// so the host decides where they live (workbook snapshot resource, sidecar
// JSON next to the xlsx, a DB row). The fidelity layer will additionally map
// specs to OOXML chart parts on save so the charts exist inside the .xlsx
// itself — this manager doesn't know or care.

export const CHART_COMPONENT_KEY = 'injoffice-chart'

interface MountedChart {
  spec: ChartSpec
  dispose: () => void
  /** The file drawing anchor the chart mounted at (file-bridge charts) —
   *  reused on save so the chart is written back where the file had it. */
  cellAnchor?: FileChartAnchor
}

/** Optional bridge for applying model order to a concrete drawing host. The
 * ids are ordered back-to-front and always belong to one worksheet. Return
 * false (or throw) to refuse the change before the model is mutated. */
export interface ChartLayerHost {
  setChartOrder(sheetId: string, chartIds: readonly string[]): boolean | void
}

export interface ChartManagerOptions {
  layerHost?: ChartLayerHost
}

let seq = 0
function newChartId(): string {
  // Time-prefixed so ids sort by creation and never collide across sessions.
  return `chart-${Date.now().toString(36)}-${(++seq).toString(36)}`
}

export class ChartManager {
  private readonly api: FUniver
  private readonly charts = new Map<string, MountedChart>()
  private commandSub: { dispose(): void } | null = null
  private refreshTimer: ReturnType<typeof setTimeout> | null = null
  private dirty: Set<string> | 'all' = new Set()
  private cascade = 0
  private theme: ChartTheme | undefined
  private readonly layerHost: ChartLayerHost | undefined
  private readonly changeListeners = new Set<() => void>()

  constructor(api: FUniver, theme?: ChartTheme, options: ChartManagerOptions = {}) {
    this.api = api
    this.theme = theme
    this.layerHost = options.layerHost
  }

  /** Host theming (dashboard light/dark). Re-renders every chart. */
  setTheme(theme: ChartTheme | undefined): void {
    this.theme = theme
    this.refreshAll()
  }

  /** Notifies on any add/remove/spec change — what a config panel binds to. */
  onChange(listener: () => void): () => void {
    this.changeListeners.add(listener)
    return () => this.changeListeners.delete(listener)
  }

  private emitChange(): void {
    this.changeListeners.forEach((l) => l())
  }

  /** Call once, after registering ChartFloat under CHART_COMPONENT_KEY. */
  start(): void {
    if (this.commandSub) return
    // One debounced refresh per command burst, scoped to charts whose source
    // intersects the edited cells. A paste/fill mutates many cells in a row;
    // selection/style commands don't change source data and are ignored.
    this.commandSub = this.api.onCommandExecuted((info) => {
      if (!this.markDirty(info)) return
      if (this.refreshTimer) clearTimeout(this.refreshTimer)
      this.refreshTimer = setTimeout(() => this.flushDirty(), 120)
    })
  }

  stop(): void {
    this.commandSub?.dispose()
    this.commandSub = null
    if (this.refreshTimer) clearTimeout(this.refreshTimer)
    this.refreshTimer = null
    this.dirty = new Set()
    for (const id of [...this.charts.keys()]) this.remove(id)
  }

  /** Insert a chart over the current selection. Returns null when there is no
   *  selection to chart. */
  createFromSelection(type: ChartType, title?: string): ChartSpec | null {
    const wb = this.api.getActiveWorkbook()
    const sheet = wb?.getActiveSheet()
    const range = sheet?.getActiveRange() ?? wb?.getActiveRange()
    if (!wb || !sheet || !range) return null
    const r = range.getRange()
    const ref: CellRangeRef = {
      sheetId: sheet.getSheetId(),
      startRow: r.startRow,
      startColumn: r.startColumn,
      endRow: r.endRow,
      endColumn: r.endColumn,
    }
    const spec: ChartSpec = { id: newChartId(), type, title, range: ref }
    return this.add(spec) ? spec : null
  }

  /** Mount a chart from a full spec (hydrate path and agent path). An
   *  optional cellAnchor (from a file's drawing anchor) places the chart over
   *  those grid cells instead of the pixel cascade — used by the file bridge
   *  so agent-made charts appear where the file says they live. */
  add(spec: ChartSpec, cellAnchor?: FileChartAnchor): boolean {
    const sheet = this.api.getActiveWorkbook()?.getSheetBySheetId(spec.range.sheetId)
    if (!sheet || this.charts.has(spec.id)) return false

    this.publish(spec)

    // Univer types float-dom data as its Serializable (indexed) shape;
    // ChartSpec is plain JSON but has no index signature, hence the cast.
    // The spec riding along in data is what makes charts come back from a
    // snapshot: hydrate() can rebuild the manager from getAllFloatDoms().
    const data = JSON.parse(JSON.stringify({ chartId: spec.id, spec }))

    let mounted: Nullable<{ id: string; dispose: () => void }>
    if (cellAnchor) {
      const range = sheet.getRange(
        cellAnchor.FromRow,
        cellAnchor.FromCol,
        Math.max(1, cellAnchor.ToRow - cellAnchor.FromRow),
        Math.max(1, cellAnchor.ToCol - cellAnchor.FromCol),
      )
      mounted = sheet.addFloatDomToRange(
        range,
        { componentKey: CHART_COMPONENT_KEY, data, allowTransform: true },
        {},
        spec.id,
      )
    }
    if (!mounted) {
      // Cascading initial placement, Excel-style: charts land in the top-left
      // working area, each subsequent one offset so none fully covers another.
      const offset = (this.cascade++ % 6) * 28
      mounted = sheet.addFloatDomToPosition(
        {
          componentKey: CHART_COMPONENT_KEY,
          initPosition: { startX: 120 + offset, endX: 580 + offset, startY: 90 + offset, endY: 380 + offset },
          data,
          allowTransform: true,
        },
        spec.id,
      )
    }
    if (!mounted) {
      dropChart(spec.id)
      return false
    }
    this.charts.set(spec.id, { spec, dispose: mounted.dispose, cellAnchor })
    this.emitChange()
    return true
  }

  remove(id: string): void {
    const chart = this.charts.get(id)
    if (!chart) return
    chart.dispose()
    this.charts.delete(id)
    dropChart(id)
    this.emitChange()
  }

  updateSpec(id: string, patch: Partial<Omit<ChartSpec, 'id'>>): void {
    const chart = this.charts.get(id)
    if (!chart) return
    chart.spec = { ...chart.spec, ...patch, id }
    this.publish(chart.spec)
    this.emitChange()
  }

  getSpec(id: string): ChartSpec | undefined {
    return this.charts.get(id)?.spec
  }

  /** Extracted series names for a chart — what the config panel's per-series
   *  override rows are keyed by. */
  getSeriesNames(id: string): string[] {
    const spec = this.charts.get(id)?.spec
    if (!spec) return []
    return extractChartData(this.readRange(spec.range), spec).series.map((s) => s.name)
  }

  list(): ChartSpec[] {
    return [...this.charts.values()].map((c) => c.spec)
  }

  /** Reorder one chart relative to charts on the same worksheet. The manager
   * updates its durable back-to-front order only after an injected visual host
   * accepts the same order. Without a host, this remains a model operation and
   * takes effect visually when the snapshot is next mounted. */
  layer(id: string, operation: ChartLayerOperation): boolean {
    const plan = this.layerPlan(id, operation)
    if (!plan) return false
    const { sheetId, sheetEntries, currentIndex, nextIndex } = plan
    const [moved] = sheetEntries.splice(currentIndex, 1)
    sheetEntries.splice(nextIndex, 0, moved)
    const chartIds = Object.freeze(sheetEntries.map(([chartId]) => chartId))
    try {
      const accepted = this.layerHost?.setChartOrder(sheetId, chartIds)
      const runtimeAccepted: unknown = accepted
      if (accepted === false || (typeof runtimeAccepted === 'object' && runtimeAccepted !== null && typeof (runtimeAccepted as { then?: unknown }).then === 'function')) return false
    } catch {
      return false
    }

    const reordered = sheetEntries[Symbol.iterator]()
    const replacement = [...this.charts.entries()].map((entry): [string, MountedChart] =>
      entry[1].spec.range.sheetId === sheetId ? reordered.next().value! : entry)
    this.charts.clear()
    for (const [chartId, mounted] of replacement) this.charts.set(chartId, mounted)
    this.emitChange()
    return true
  }

  canLayer(id: string, operation: ChartLayerOperation): boolean {
    return this.layerPlan(id, operation) !== null
  }

  private layerPlan(id: string, operation: ChartLayerOperation): {
    sheetId: string
    sheetEntries: Array<[string, MountedChart]>
    currentIndex: number
    nextIndex: number
  } | null {
    const chart = this.charts.get(id)
    if (!chart) return null
    const sheetId = chart.spec.range.sheetId
    const sheetEntries = [...this.charts.entries()].filter(([, mounted]) => mounted.spec.range.sheetId === sheetId)
    const currentIndex = sheetEntries.findIndex(([chartId]) => chartId === id)
    let nextIndex = currentIndex
    if (operation === 'bringForward') nextIndex = Math.min(sheetEntries.length - 1, currentIndex + 1)
    else if (operation === 'sendBackward') nextIndex = Math.max(0, currentIndex - 1)
    else if (operation === 'bringToFront') nextIndex = sheetEntries.length - 1
    else if (operation === 'sendToBack') nextIndex = 0
    else return null
    return nextIndex === currentIndex ? null : { sheetId, sheetEntries, currentIndex, nextIndex }
  }

  serialize(): ChartSpec[] {
    return this.list()
  }

  /** Specs plus their file anchors (when known) — what the save-time OOXML
   *  writer consumes (toFile.ts). */
  listWithAnchors(): Array<{ spec: ChartSpec; cellAnchor?: FileChartAnchor }> {
    return [...this.charts.values()].map((c) => ({ spec: c.spec, cellAnchor: c.cellAnchor }))
  }

  hydrate(specs: ChartSpec[]): void {
    for (const spec of specs) this.add(spec)
  }

  private markDirty(info: { id?: string; type?: number; params?: unknown }): boolean {
    const mode = invalidateMode(info)
    if (mode.kind === 'skip') return false
    if (mode.kind === 'all' || this.dirty === 'all') {
      this.dirty = 'all'
      return true
    }
    const dirty = this.dirty
    for (const [id, chart] of this.charts) {
      if (sourceNeedsRefresh(chart.spec.range, mode)) dirty.add(id)
    }
    return dirty.size > 0
  }

  private flushDirty(): void {
    const dirty = this.dirty
    this.dirty = new Set()
    if (dirty === 'all') {
      this.refreshAll()
      return
    }
    for (const id of dirty) {
      const chart = this.charts.get(id)
      if (chart) this.publish(chart.spec)
    }
  }

  private refreshAll(): void {
    for (const chart of this.charts.values()) this.publish(chart.spec)
  }

  private publish(spec: ChartSpec): void {
    const grid = this.readRange(spec.range)
    const data = extractChartData(grid, spec)
    publishChartOption(spec.id, buildEChartsOption(spec, data, this.theme))
  }

  private readRange(ref: CellRangeRef): unknown[][] {
    const sheet = this.api.getActiveWorkbook()?.getSheetBySheetId(ref.sheetId)
    if (!sheet) return []
    const rows = ref.endRow - ref.startRow + 1
    const cols = ref.endColumn - ref.startColumn + 1
    if (rows <= 0 || cols <= 0) return []
    const range = sheet.getRange(ref.startRow, ref.startColumn, rows, cols)
    // Univer's getValues() includes display formats (e.g. "$1,250.00").
    // Charts need the underlying numeric values, with a fallback for older hosts.
    return typeof range.getRawValues === 'function' ? range.getRawValues() : range.getValues()
  }
}
