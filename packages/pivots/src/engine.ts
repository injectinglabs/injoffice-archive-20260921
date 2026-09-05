import type { AggKind, BakedPivot, PivotFieldMember, PivotFieldSort, PivotSpec } from './types'

// The pivot engine: pure aggregation from a raw values grid to a baked grid.
// No Univer, no DOM — fully unit-tested, and the same function the future
// pivot.create RPC and the OOXML writer's cached-grid path will call.
//
// Semantics (v1, chosen to match what Excel shows for the same config):
//  - First source row = field names. Duplicate field names keep the first.
//  - Row grouping: one baked column per row field; groups ordered by first
//    appearance in the source (stable, predictable for users and tests).
//  - Column grouping: 0 or 1 field. With a column field, one baked column per
//    (column value × value field); without, one per value field.
//  - Aggregations: sum/avg/min/max coerce numerically (non-numeric cells are
//    ignored); count counts non-empty cells.
//  - Filters: field → allowed values, applied to source rows before grouping.
//  - Grand total row at the bottom (default on) aggregating the FILTERED
//    source rows.

function toNumber(v: unknown): number | null {
  if (typeof v === 'number' && Number.isFinite(v)) return v
  if (typeof v === 'string') {
    const t = v.trim().replace(/,/g, '')
    if (t === '') return null
    const n = Number(t)
    if (Number.isFinite(n)) return n
  }
  return null
}

function cellKey(v: unknown): string {
  if (v == null) return ''
  return String(v).trim()
}

function aggregate(kind: AggKind, cells: unknown[]): number | null {
  if (kind === 'count') {
    return cells.filter((c) => c != null && String(c).trim() !== '').length
  }
  const nums = cells.map(toNumber).filter((n): n is number => n !== null)
  if (nums.length === 0) return null
  switch (kind) {
    case 'sum':
      return nums.reduce((a, b) => a + b, 0)
    case 'avg':
      return nums.reduce((a, b) => a + b, 0) / nums.length
    case 'min':
      return Math.min(...nums)
    case 'max':
      return Math.max(...nums)
  }
}

const KEY_SEP = '\0'

export function bakePivot(
  sourceGrid: unknown[][],
  spec: Pick<PivotSpec, 'rows' | 'columns' | 'values' | 'filters' | 'memberFilters' | 'pageFields' | 'sorts' | 'grandTotals'>,
): BakedPivot {
  if (sourceGrid.length === 0) return { grid: [], rowCount: 0, columnCount: 0 }
  const header = sourceGrid[0].map(cellKey)
  const fieldIndex = new Map<string, number>()
  header.forEach((name, i) => {
    if (name && !fieldIndex.has(name)) fieldIndex.set(name, i)
  })

  const rowFields = spec.rows.filter((f) => fieldIndex.has(f))
  const colField = spec.columns.find((f) => fieldIndex.has(f))
  const valueFields = spec.values.filter((v) => fieldIndex.has(v.field))
  if (valueFields.length === 0) return { grid: [], rowCount: 0, columnCount: 0 }

  // Filter source rows.
  let body = sourceGrid.slice(1)
  if (spec.filters) {
    for (const [field, allowed] of Object.entries(spec.filters)) {
      const idx = fieldIndex.get(field)
      if (idx === undefined) continue
      const set = new Set(allowed)
      body = body.filter((r) => set.has(cellKey(r[idx])))
    }
  }
  for (const filter of spec.memberFilters ?? []) {
    const idx = fieldIndex.get(filter.field)
    if (idx === undefined) continue
    const values = new Set(filter.values)
    body = body.filter((row) => filter.mode === 'include'
      ? values.has(cellKey(row[idx]))
      : !values.has(cellKey(row[idx])))
  }
  for (const page of spec.pageFields ?? []) {
    if (page.selectedItem === undefined) continue
    const idx = fieldIndex.get(page.field)
    if (idx === undefined) continue
    body = body.filter((row) => cellKey(row[idx]) === page.selectedItem)
  }

  const sortByField = new Map<string, PivotFieldSort['direction']>()
  for (const rule of spec.sorts ?? []) if (!sortByField.has(rule.field)) sortByField.set(rule.field, rule.direction)

  // Column group values in first-appearance order.
  const colValues: string[] = []
  if (colField !== undefined) {
    const seen = new Set<string>()
    const ci = fieldIndex.get(colField)!
    for (const r of body) {
      const v = cellKey(r[ci])
      if (!seen.has(v)) {
        seen.add(v)
        colValues.push(v)
      }
    }
    const direction = sortByField.get(colField)
    if (direction) colValues.sort((left, right) => compareLabel(left, right, direction))
  }

  // Group rows by composite row key, first-appearance order.
  const groups = new Map<string, { labels: string[]; rows: unknown[][] }>()
  const order: string[] = []
  for (const r of body) {
    const labels = rowFields.map((f) => cellKey(r[fieldIndex.get(f)!]))
    const key = labels.join(KEY_SEP)
    let g = groups.get(key)
    if (!g) {
      g = { labels, rows: [] }
      groups.set(key, g)
      order.push(key)
    }
    g.rows.push(r)
  }
  order.sort((left, right) => {
    const leftLabels = groups.get(left)!.labels
    const rightLabels = groups.get(right)!.labels
    for (let index = 0; index < rowFields.length; index++) {
      if (leftLabels[index] === rightLabels[index]) continue
      const direction = sortByField.get(rowFields[index])
      return direction ? compareLabel(leftLabels[index], rightLabels[index], direction) : 0
    }
    return 0
  })

  const valueLabel = (v: { field: string; agg: AggKind }): string =>
    `${v.agg === 'avg' ? 'Average' : v.agg[0].toUpperCase() + v.agg.slice(1)} of ${v.field}`

  // Header row of the baked grid: row-field names (or one blank label column
  // when there are no row fields), then per column-group×value columns.
  const labelCols = Math.max(rowFields.length, 1)
  const headerRow: (string | number | null)[] = rowFields.length > 0 ? [...rowFields] : ['']
  const dataCols: { colValue: string | null; vf: (typeof valueFields)[number] }[] = []
  if (colField !== undefined) {
    for (const cv of colValues) {
      for (const vf of valueFields) {
        dataCols.push({ colValue: cv, vf })
        headerRow.push(
          valueFields.length > 1 ? `${cv} — ${valueLabel(vf)}` : `${cv || '(blank)'}`,
        )
      }
    }
    // Row totals across column groups (per value field) when a column field
    // is present.
    for (const vf of valueFields) {
      dataCols.push({ colValue: null, vf })
      headerRow.push(`Total ${valueLabel(vf)}`)
    }
  } else {
    for (const vf of valueFields) {
      dataCols.push({ colValue: null, vf })
      headerRow.push(valueLabel(vf))
    }
  }

  const grid: (string | number | null)[][] = [headerRow]
  const colIdx = colField !== undefined ? fieldIndex.get(colField)! : -1

  const rowFor = (rows: unknown[][], labels: (string | number | null)[]) => {
    const out: (string | number | null)[] = [...labels]
    for (const dc of dataCols) {
      const subset =
        dc.colValue === null ? rows : rows.filter((r) => cellKey(r[colIdx]) === dc.colValue)
      const vi = fieldIndex.get(dc.vf.field)!
      out.push(aggregate(dc.vf.agg, subset.map((r) => r[vi])))
    }
    return out
  }

  if (rowFields.length > 0) {
    for (const key of order) {
      const g = groups.get(key)!
      grid.push(rowFor(g.rows, g.labels))
    }
    if (spec.grandTotals !== false) {
      const totalLabels: (string | number | null)[] = rowFields.map((_, i) =>
        i === 0 ? 'Grand Total' : '',
      )
      grid.push(rowFor(body, totalLabels))
    }
  } else {
    // No row fields: the pivot is a single summary row — a separate grand
    // total would just duplicate it.
    grid.push(rowFor(body, ['Total']))
  }

  const width = labelCols + dataCols.length
  for (const r of grid) while (r.length < width) r.push(null)

  return { grid, rowCount: grid.length, columnCount: width }
}

function compareLabel(left: string, right: string, direction: PivotFieldSort['direction']): number {
  const order = left < right ? -1 : left > right ? 1 : 0
  return direction === 'descending' ? -order : order
}

/** Distinct values of one field in the source (for slicer chips), in
 *  first-appearance order, UNfiltered — a slicer must show deselected values. */
export function fieldValues(sourceGrid: unknown[][], field: string): string[] {
  if (sourceGrid.length === 0) return []
  const idx = sourceGrid[0].map(cellKey).indexOf(field)
  if (idx < 0) return []
  const seen = new Set<string>()
  const out: string[] = []
  for (const r of sourceGrid.slice(1)) {
    const v = cellKey(r[idx])
    if (!seen.has(v)) {
      seen.add(v)
      out.push(v)
    }
  }
  return out
}

/** Complete, stable member inventory used by native pivot conversion. Values
 * whose runtime type OOXML sharedItems cannot safely represent are retained as
 * `unsupported`, so the adapter refuses instead of changing filter meaning. */
export function fieldMembers(sourceGrid: unknown[][], field: string): PivotFieldMember[] {
  if (sourceGrid.length === 0) return []
  const idx = sourceGrid[0].map(cellKey).indexOf(field)
  if (idx < 0) return []
  const seen = new Set<string>()
  const out: PivotFieldMember[] = []
  for (const row of sourceGrid.slice(1)) {
    const raw = row[idx]
    const value = cellKey(raw)
    const kind: PivotFieldMember['kind'] = raw == null || value === ''
      ? 'blank'
      : typeof raw === 'string'
        ? 'string'
        : typeof raw === 'number' && Number.isFinite(raw)
          ? 'number'
          : typeof raw === 'boolean'
            ? 'boolean'
            : 'unsupported'
    const key = `${kind}:${value}`
    if (seen.has(key)) continue
    seen.add(key)
    out.push({ value, kind })
  }
  return out
}

/** Field names available in a source grid (header row, deduped). */
export function sourceFields(sourceGrid: unknown[][]): string[] {
  if (sourceGrid.length === 0) return []
  const seen = new Set<string>()
  const out: string[] = []
  for (const name of sourceGrid[0].map(cellKey)) {
    if (name && !seen.has(name)) {
      seen.add(name)
      out.push(name)
    }
  }
  return out
}
