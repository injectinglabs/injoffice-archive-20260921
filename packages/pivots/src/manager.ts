import type { ICellData } from '@univerjs/core'
import type { FUniver } from '@univerjs/core/lib/facade'
// Facade surface arrives via module augmentation (same pattern as
// @injoffice/charts — see that manager's note).
import type {} from '@univerjs/sheets/lib/facade'
import type {} from '@univerjs/sheets-ui/lib/facade'
import type {} from '@univerjs/ui/lib/facade'
import { bakePivot, fieldValues, sourceFields } from './engine'
import { invalidateMode, rangesIntersect, sourceNeedsRefresh } from './invalidate'
import { toWirePivotRemove, toWirePivotUpdate, type PivotLifecycleWireResult, type PivotWireContext, type WirePivotRemove, type WirePivotUpdate } from './toFile'
import type { CellRangeRef, PivotSpec } from './types'

// PivotManager — the host-facing API of @injoffice/pivots.
//
// Owns the spec list and keeps every pivot's BAKED GRID written into the
// sheet at its target, re-baking (debounced) when the edited cells intersect
// that pivot's source — not on every command, and never because the bake
// itself wrote the target. The baked grid is real cells — it exports, copies,
// and charts like any other data; the spec is what makes it a live, EDITABLE
// pivot rather than a one-shot paste (the panel mutates the spec, the
// manager re-bakes).
//
// Self-write guard: baking writes cells, which fires the same command events
// we listen to — `bakingDepth` keeps that from looping. Nested clearExtent
// must not drop the outer guard. After a bake writes a target, other pivots
// whose source intersects that target (pivot-of-pivot) are rebaked in the
// same turn; the in-flight set bounds cycles.

let seq = 0
function newPivotId(): string {
  return `pivot-${Date.now().toString(36)}-${(++seq).toString(36)}`
}

interface MountedPivot {
  spec: PivotSpec
  /** Extent of the last bake (rows/cols) so a smaller re-bake clears leftovers. */
  lastRows: number
  lastCols: number
}

export class PivotManager {
  private readonly api: FUniver
  private readonly pivots = new Map<string, MountedPivot>()
  private commandSub: { dispose(): void } | null = null
  private refreshTimer: ReturnType<typeof setTimeout> | null = null
  private dirty: Set<string> | 'all' = new Set()
  private bakingDepth = 0
  private readonly inFlight = new Set<string>()
  private readonly changeListeners = new Set<() => void>()

  constructor(api: FUniver) {
    this.api = api
  }

  start(): void {
    if (this.commandSub) return
    this.commandSub = this.api.onCommandExecuted((info, options) => {
      // Collaboration: a remote peer's edits (fromCollab) arrive as data,
      // including THEIR baked pivot cells. Rebaking here would emit local
      // mutations, ship them, and make the peer rebake in turn (ping-pong).
      if (this.bakingDepth > 0 || (options as { fromCollab?: boolean } | undefined)?.fromCollab) return
      if (!this.markDirty(info)) return
      if (this.refreshTimer) clearTimeout(this.refreshTimer)
      this.refreshTimer = setTimeout(() => this.flushDirty(), 200)
    })
  }

  stop(): void {
    this.commandSub?.dispose()
    this.commandSub = null
    if (this.refreshTimer) clearTimeout(this.refreshTimer)
    this.refreshTimer = null
    this.dirty = new Set()
    this.bakingDepth = 0
    this.inFlight.clear()
    this.pivots.clear()
  }

  onChange(listener: () => void): () => void {
    this.changeListeners.add(listener)
    return () => this.changeListeners.delete(listener)
  }

  private emitChange(): void {
    this.changeListeners.forEach((l) => l())
  }

  /** Create a pivot from the current selection (selection = source block
   *  including headers). Default config: first field as rows, last field
   *  summed; target lands two columns right of the source. */
  createFromSelection(): PivotSpec | null {
    const wb = this.api.getActiveWorkbook()
    const sheet = wb?.getActiveSheet()
    const range = sheet?.getActiveRange() ?? wb?.getActiveRange()
    if (!wb || !sheet || !range) return null
    const r = range.getRange()
    const source: CellRangeRef = {
      sheetId: sheet.getSheetId(),
      startRow: r.startRow,
      startColumn: r.startColumn,
      endRow: r.endRow,
      endColumn: r.endColumn,
    }
    const grid = this.readRange(source)
    const fields = sourceFields(grid)
    if (fields.length < 2) return null
    const spec: PivotSpec = {
      id: newPivotId(),
      source,
      rows: [fields[0]],
      columns: [],
      values: [{ field: fields[fields.length - 1], agg: 'sum' }],
      target: {
        sheetId: source.sheetId,
        startRow: source.startRow,
        startColumn: source.endColumn + 2,
      },
    }
    return this.add(spec) ? spec : null
  }

  add(spec: PivotSpec): boolean {
    if (this.pivots.has(spec.id)) return false
    this.pivots.set(spec.id, { spec, lastRows: 0, lastCols: 0 })
    this.bake(spec.id)
    this.emitChange()
    return true
  }

  /** Remove the pivot and blank its baked cells. */
  remove(id: string): void {
    const p = this.pivots.get(id)
    if (!p) return
    this.clearExtent(p)
    this.pivots.delete(id)
    this.emitChange()
  }

  updateSpec(id: string, patch: Partial<Omit<PivotSpec, 'id'>>): void {
    const p = this.pivots.get(id)
    if (!p) return
    p.spec = { ...p.spec, ...patch, id }
    this.bake(id)
    this.emitChange()
  }

  getSpec(id: string): PivotSpec | undefined {
    return this.pivots.get(id)?.spec
  }

  list(): PivotSpec[] {
    return [...this.pivots.values()].map((p) => p.spec)
  }

  /** Preflight used by host command/snapshot adapters. It does not mutate. */
  canMount(spec: Pick<PivotSpec, 'source' | 'target'>): boolean {
    const workbook = this.api.getActiveWorkbook()
    return !!workbook?.getSheetBySheetId(spec.source.sheetId)
      && !!workbook.getSheetBySheetId(spec.target.sheetId)
  }

  /** Source headers for an unmounted candidate spec. */
  sourceFieldsAt(source: CellRangeRef): string[] {
    return sourceFields(this.readRange(source))
  }

  serialize(): PivotSpec[] {
    return this.list()
  }

  /** Build a native delete request without removing the baked grid first. The
   * caller applies the file operation, then calls remove after it succeeds. */
  nativeRemoveRequest(id: string): PivotLifecycleWireResult<WirePivotRemove> {
    const spec = this.pivots.get(id)?.spec
    if (!spec) return { skipped: [`pivot ${id}: it is not managed`] }
    return toWirePivotRemove(spec)
  }

  /** Build an identity-bound native update for the manager's current spec. */
  nativeUpdateRequest(id: string, context: PivotWireContext): PivotLifecycleWireResult<WirePivotUpdate> {
    const spec = this.pivots.get(id)?.spec
    if (!spec) return { skipped: [`pivot ${id}: it is not managed`] }
    return toWirePivotUpdate(spec, context)
  }

  hydrate(specs: PivotSpec[]): void {
    for (const spec of specs) this.add(spec)
  }

  /** Fields available in a pivot's source (for the panel). */
  sourceFieldsFor(id: string): string[] {
    const p = this.pivots.get(id)
    if (!p) return []
    return sourceFields(this.readRange(p.spec.source))
  }

  /** Distinct values of a field (for slicer chips), unfiltered. */
  fieldValuesFor(id: string, field: string): string[] {
    const p = this.pivots.get(id)
    if (!p) return []
    return fieldValues(this.readRange(p.spec.source), field)
  }

  private markDirty(info: { id?: string; type?: number; params?: unknown }): boolean {
    const mode = invalidateMode(info)
    if (mode.kind === 'skip') return false
    if (mode.kind === 'all' || this.dirty === 'all') {
      this.dirty = 'all'
      return true
    }
    const dirty = this.dirty
    for (const [id, pivot] of this.pivots) {
      if (sourceNeedsRefresh(pivot.spec.source, mode)) dirty.add(id)
    }
    return dirty.size > 0
  }

  private flushDirty(): void {
    const dirty = this.dirty
    this.dirty = new Set()
    if (dirty === 'all') {
      this.rebakeAll()
      return
    }
    for (const id of dirty) this.bake(id)
  }

  private rebakeAll(): void {
    for (const id of this.pivots.keys()) this.bake(id)
  }

  private bake(id: string): void {
    if (this.inFlight.has(id)) return
    const p = this.pivots.get(id)
    if (!p) return
    const grid = this.readRange(p.spec.source)
    const baked = bakePivot(grid, p.spec)
    const sheet = this.api.getActiveWorkbook()?.getSheetBySheetId(p.spec.target.sheetId)
    if (!sheet) return
    this.inFlight.add(id)
    const previous = this.extentOf(p)
    this.bakingDepth++
    try {
      // Blank any cells the previous (larger) bake covered, then write.
      if (p.lastRows > baked.rowCount || p.lastCols > baked.columnCount) {
        this.clearExtent(p)
      }
      if (baked.rowCount > 0) {
        // ICellData form so null cells BLANK their targets (plain CellValue
        // arrays don't admit null).
        const cells: ICellData[][] = baked.grid.map((row) => row.map((v) => ({ v })))
        sheet
          .getRange(p.spec.target.startRow, p.spec.target.startColumn, baked.rowCount, baked.columnCount)
          .setValues(cells)
      }
      p.lastRows = baked.rowCount
      p.lastCols = baked.columnCount
      this.bakeDependents(id, previous, this.extentOf(p))
    } finally {
      this.bakingDepth--
      this.inFlight.delete(id)
    }
  }

  private bakeDependents(bakedId: string, ...extents: Array<CellRangeRef | null>): void {
    for (const [id, pivot] of this.pivots) {
      if (id === bakedId) continue
      if (extents.some((extent) => extent && rangesIntersect(pivot.spec.source, extent))) this.bake(id)
    }
  }

  private extentOf(p: MountedPivot): CellRangeRef | null {
    if (p.lastRows <= 0 || p.lastCols <= 0) return null
    const t = p.spec.target
    return {
      sheetId: t.sheetId,
      startRow: t.startRow,
      startColumn: t.startColumn,
      endRow: t.startRow + p.lastRows - 1,
      endColumn: t.startColumn + p.lastCols - 1,
    }
  }

  private clearExtent(p: MountedPivot): void {
    if (p.lastRows === 0 || p.lastCols === 0) return
    const sheet = this.api.getActiveWorkbook()?.getSheetBySheetId(p.spec.target.sheetId)
    if (!sheet) return
    const blanks: ICellData[][] = Array.from({ length: p.lastRows }, () =>
      Array.from({ length: p.lastCols }, () => ({ v: null })),
    )
    sheet
      .getRange(p.spec.target.startRow, p.spec.target.startColumn, p.lastRows, p.lastCols)
      .setValues(blanks)
  }

  private readRange(ref: CellRangeRef): unknown[][] {
    const sheet = this.api.getActiveWorkbook()?.getSheetBySheetId(ref.sheetId)
    if (!sheet) return []
    const rows = ref.endRow - ref.startRow + 1
    const cols = ref.endColumn - ref.startColumn + 1
    if (rows <= 0 || cols <= 0) return []
    const range = sheet.getRange(ref.startRow, ref.startColumn, rows, cols)
    // Aggregate numbers, not currency/percentage display strings.
    return typeof range.getRawValues === 'function' ? range.getRawValues() : range.getValues()
  }
}
