import { assessPivotRepresentability } from '../../../packages/pivots/src/representability'
import { toWirePivots } from '../../../packages/pivots/src/toFile'
import type { AggKind, PivotSpec } from '../../../packages/pivots/src/types'

export function playgroundPivotSpec(input: {
  rows: string[]
  columns: string[]
  valueField: string
  aggregation: AggKind
  filters?: Record<string, string[]>
  rowCount: number
  columnCount: number
}): PivotSpec {
  return {
    id: 'playground-sales',
    source: { sheetId: 'sales', startRow: 0, startColumn: 0, endRow: input.rowCount - 1, endColumn: input.columnCount - 1 },
    rows: input.rows,
    columns: input.columns,
    values: [{ field: input.valueField, agg: input.aggregation }],
    target: { sheetId: 'sales', startRow: 0, startColumn: input.columnCount + 2 },
    filters: input.filters,
    grandTotals: true,
  }
}

export function playgroundPivotWire(spec: PivotSpec, headerRow: unknown[], members: Record<string, string[]> = {}) {
  const fields = headerRow.map((cell) => String(cell ?? ''))
  return {
    representability: assessPivotRepresentability(spec, {
      fields,
      membersOf: (field) => (members[field] ?? []).map((value) => ({ value, kind: 'string' as const })),
    }),
    wire: toWirePivots([spec], {
      sheetNameOf: (sheetId) => sheetId === 'sales' ? 'Sales' : null,
      headerRowOf: () => headerRow,
    }),
  }
}
