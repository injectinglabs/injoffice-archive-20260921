// Pivot model for @injoffice/pivots.
//
// Same contract philosophy as ChartSpec: PivotSpec is plain JSON, independent
// of Univer — it's what agents send over a pivot.create RPC, what persists,
// what the editing panel mutates, and what the Go fidelity layer will map to
// native OOXML pivot parts (pivotCache + pivotTable with refreshOnLoad=1, so
// Excel rebuilds a live pivot on open).
//
// The spec is EDITABLE by design: the panel mutates the spec and the engine
// re-bakes.

/** A rectangular block of cells on one sheet, 0-indexed, inclusive.
 *  (Same shape as @injoffice/charts' CellRangeRef; kept local so the two
 *  packages stay independently consumable.) */
export interface CellRangeRef {
  sheetId: string
  startRow: number
  startColumn: number
  endRow: number
  endColumn: number
}

export type AggKind = 'sum' | 'count' | 'avg' | 'min' | 'max'

export const AGG_KINDS: readonly AggKind[] = ['sum', 'count', 'avg', 'min', 'max']

export interface PivotValueField {
  /** Source column header this value aggregates. */
  field: string
  agg: AggKind
}

/** A page/report-filter field. Omit selectedItem to show all members. */
export interface PivotPageField {
  field: string
  selectedItem?: string
}

/** Member-level filtering. Values use the source cell's display-neutral string form. */
export interface PivotMemberFilter {
  field: string
  mode: 'include' | 'exclude'
  values: string[]
}

/** Label ordering only. Value-based/autosort scopes are intentionally excluded. */
export interface PivotFieldSort {
  field: string
  direction: 'ascending' | 'descending'
}

export type PivotMemberKind = 'string' | 'number' | 'boolean' | 'blank' | 'unsupported'

export interface PivotFieldMember {
  value: string
  kind: PivotMemberKind
}

/** Stable identity returned by native XLSX pivot hydration. Unlike a display
 * name or relationship id, the pivot-table part names exactly one package
 * object and is retained by a native update. */
export interface NativePivotIdentity {
  part: string
}

export interface PivotSpec {
  id: string
  /** Present only when this spec was hydrated from a native XLSX pivot. */
  nativeIdentity?: NativePivotIdentity
  /** Source data block INCLUDING the header row (field names). */
  source: CellRangeRef
  /** Row group fields, outermost first. Multi-field = one column per field
   *  in the baked grid (flat composite grouping, not collapsible trees). */
  rows: string[]
  /** Optional single column group field (v1: 0 or 1 — matches what the OOXML
   *  writer will emit first). */
  columns: string[]
  values: PivotValueField[]
  /** Slicer state: field → allowed values. Absent field = no filter. Empty
   *  array = filter everything (an explicit "nothing selected" state). */
  filters?: Record<string, string[]>
  /** Explicit member filters used by the renderer and bounded native XLSX adapter. */
  memberFilters?: PivotMemberFilter[]
  /** Page/report filters, in display order. */
  pageFields?: PivotPageField[]
  /** Ascending/descending label sorts for row or column fields. */
  sorts?: PivotFieldSort[]
  /** Top-left corner where the baked grid lands. */
  target: { sheetId: string; startRow: number; startColumn: number }
  /** Include a grand-total row (default true). */
  grandTotals?: boolean
}

/** The engine's output: a rectangular values grid (headers included) plus its
 *  dimensions — what the manager writes to the target range and what the
 *  OOXML writer will bake alongside the native parts. */
export interface BakedPivot {
  grid: (string | number | null)[][]
  rowCount: number
  columnCount: number
}
