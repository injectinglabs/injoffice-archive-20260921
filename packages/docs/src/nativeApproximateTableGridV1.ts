/** Approximate-preview-only sizing of auto-width tables from their authored
 * tblGrid. Strict page paint never calls this; it is a declared policy of the
 * read-only current-layout approximate preview. No source bytes change. */
import { nativeDocxCellWidthAgreesWithGridV1 } from './nativeContract.js'
import type { NativeDocxTableV1 } from './nativeContract.js'

export const DOCX_APPROXIMATE_TABLE_GRID_POLICY = 'approximate-authored-grid-fitted-v1' as const

export interface NativeDocxApproximateTableGridPolicyV1 {
  name: typeof DOCX_APPROXIMATE_TABLE_GRID_POLICY
  section_id: string
  container_width_twips: number
  /** Container minus the table indent: the column width the fitted grid sums to. */
  available_width_twips: number
  source_grid_widths_twips: number[]
  source_cell_widths_twips: (number | null)[][]
  fitted_grid_widths_twips: number[]
}

const LIMIT = 20_000_000
const bounded = (n: unknown): n is number => typeof n === 'number' && Number.isSafeInteger(n) && n >= 0 && n <= LIMIT

/**
 * Applies only when today's content autofit would otherwise collapse a table,
 * or the approximate fixed-width fallback would paint it past its column,
 * although its consistent authored preferences do not fit the text column:
 * layout is autofit (explicit, or the cascade default of an absent
 * w:tblLayout), preferred width is auto, every cell repeats its grid slice,
 * and the grid sum exceeds the available column width.
 * Word's gridCol widths include the cell margins, so a body-wide table authored
 * by Word carries a grid wider than the column by exactly those margins; the
 * preview scales the grid proportionally to the column with deterministic
 * largest-remainder integer twips because pagination refuses any table edge
 * past the column. The fitted table then behaves exactly like an authored
 * fixed table: shaped content minima are not compared against the fitted
 * slices, so long unbreakable words wrap or overflow at the fitted width as
 * they would in a fixed grid. Anything else returns undefined so the existing
 * policies and refusals decide.
 */
export function fitNativeDocxApproximateTableGridV1(table: NativeDocxTableV1, containerWidth: number, sectionID: string): { table: NativeDocxTableV1; policy: NativeDocxApproximateTableGridPolicyV1 } | undefined {
  const grid = table.grid_widths_twips, margins = table.cell_margins
  if ((table.layout !== undefined && table.layout !== 'autofit') || table.width_twips !== undefined || table.width_percent_fiftieths !== undefined || table.alignment !== 'left') return undefined
  if (!grid?.length || grid.length > 256 || grid.some((n) => !bounded(n) || n === 0) || !margins || !bounded(containerWidth) || !bounded(table.indent_twips) || ![margins.left_twips, margins.right_twips].every(bounded)) return undefined
  if (table.rows.length === 0 || table.rows.length > 10_000) return undefined
  const gridSum = grid.reduce((a, b) => a + b, 0)
  const available = containerWidth - table.indent_twips
  if (!bounded(gridSum) || available <= 0 || gridSum <= available) return undefined
  let cells = 0
  for (const row of table.rows) {
    let column = 0
    for (const cell of row.cells) {
      if (++cells > 100_000 || !Number.isSafeInteger(cell.grid_span) || cell.grid_span < 1 || column + cell.grid_span > grid.length || cell.vertical_merge !== 'none') return undefined
      const end = column + cell.grid_span
      if (!nativeDocxCellWidthAgreesWithGridV1(cell, grid.slice(column, end).reduce((a, b) => a + b, 0))) return undefined
      column = end
    }
    if (column !== grid.length) return undefined
  }
  const remainders = grid.map((width, index) => { const numerator = BigInt(width) * BigInt(available); return { index, whole: Number(numerator / BigInt(gridSum)), remainder: numerator % BigInt(gridSum) } })
  const fitted = remainders.map((entry) => entry.whole)
  remainders.sort((a, b) => a.remainder === b.remainder ? a.index - b.index : a.remainder > b.remainder ? -1 : 1)
  const left = available - fitted.reduce((a, b) => a + b, 0)
  for (let i = 0; i < left; i += 1) fitted[remainders[i]!.index]! += 1
  if (fitted.some((n) => n <= 0)) return undefined
  const width = fitted.reduce((a, b) => a + b, 0)
  if (width !== available) return undefined
  const rows = table.rows.map((row) => {
    let column = 0
    return { ...row, cells: row.cells.map((cell) => { const end = column + cell.grid_span; const projected = { ...cell, width_twips: fitted.slice(column, end).reduce((a, b) => a + b, 0) }; column = end; return projected }) }
  })
  const policy: NativeDocxApproximateTableGridPolicyV1 = {
    name: DOCX_APPROXIMATE_TABLE_GRID_POLICY, section_id: sectionID, container_width_twips: containerWidth, available_width_twips: available,
    source_grid_widths_twips: [...grid], source_cell_widths_twips: table.rows.map((row) => row.cells.map((cell) => cell.width_twips ?? null)),
    fitted_grid_widths_twips: [...fitted],
  }
  return { policy, table: { ...table, layout: 'fixed', width_twips: width, grid_widths_twips: fitted, rows } }
}
