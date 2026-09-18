/** Approximate-preview-only sizing of percentage-width tables. Strict page
 * paint never calls this; it is a declared policy of the read-only
 * current-layout approximate preview. No source bytes change. */
import type { NativeDocxTableV1 } from './nativeContract.js'

export const DOCX_APPROXIMATE_PERCENT_TABLE_POLICY = 'approximate-percent-authored-grid-v1' as const

export interface NativeDocxApproximatePercentTablePolicyV1 {
  name: typeof DOCX_APPROXIMATE_PERCENT_TABLE_POLICY
  section_id: string
  container_width_twips: number
  percent_fiftieths: number
  /** floor(container x percent / 5000) -- the width the percentage states,
   * recorded for the reader and deliberately NOT painted. */
  percent_width_twips: number
  /** True when the percentage resolves to a whole twip, so the recorded width
   * is the exact one rather than Word's own truncation of a fraction. */
  percent_width_exact: boolean
  source_grid_widths_twips: number[]
  source_cell_widths_twips: (number | null)[][]
  /** The authored tblGrid, painted as a fixed grid, and restated on every cell. */
  painted_grid_widths_twips: number[]
  painted_width_twips: number
}

const LIMIT = 20_000_000
const bounded = (n: unknown): n is number => typeof n === 'number' && Number.isSafeInteger(n) && n >= 0 && n <= LIMIT

/**
 * Paints a percentage-width table at its authored w:tblGrid instead of
 * resolving the percentage, which is what Microsoft Word itself does.
 *
 * Measured on Word's own PDF exports of the only two percentage tables in the
 * hard-v2 corpus, on the 1/300 in grid those exports quantize to:
 *
 * - `lvlPicBulletId.docx`: w:tblW 4850 pct in a 9360 twip text column, so the
 *   percentage is 9079.2 twips. Word's authored gridCol is 9079 -- Word
 *   truncated its own resolution -- and Word paints the table 454.0 pt wide
 *   (9080 twips) from the left margin.
 * - `tdf135943_shapeWithText_LayoutInCell0_compat15.docx`: w:tblW 5000 pct in a
 *   6123 twip text column, with an authored grid of 3006 + 3111 = 6117 that a
 *   full-width resolution would have to stretch to 3008.95 + 3114.05. Word does
 *   not stretch it: it puts the vertical rule between the two cells at exactly
 *   207.00 pt = the 56.7 pt left margin + 150.30 pt = 3006 twips, and the right
 *   edge at 362.52 pt against 362.55 pt for 6117 twips -- where the stretched
 *   grid would put them at 207.15 pt and 362.85 pt, more than a quantum away.
 *
 * So the authored grid is the geometry, not a preference to be rescaled, and
 * this policy reproduces Word without any new arithmetic: no rounding of the
 * percentage is performed, because the percentage is not what Word paints.
 *
 * Two facts become approximate, both recorded in the returned policy:
 *   1. the stated percentage is not resolved -- a document whose grid does not
 *      already agree with it is painted at the grid, not at the percentage;
 *   2. a w:tcW that disagrees with the cell's grid slice is restated as that
 *      slice, because a percentage table's cell preferences cannot all hold at
 *      once and Word paints the grid (tdf135943 states 3007/3115 against
 *      3006/3111).
 *
 * Anything else returns undefined so the existing exact policy and its refusal
 * decide: in particular a grid wider than the text column, which has no lawful
 * placement and is not a shape either of these two files has.
 */
export function fitNativeDocxApproximatePercentTableV1(table: NativeDocxTableV1, containerWidth: number, sectionID: string): { table: NativeDocxTableV1; policy: NativeDocxApproximatePercentTablePolicyV1 } | undefined {
  const percent = table.width_percent_fiftieths, grid = table.grid_widths_twips
  if (table.width_twips !== undefined || !bounded(percent) || percent < 1 || percent > 5000 || !bounded(containerWidth) || containerWidth <= 0) return undefined
  if (!grid?.length || grid.length > 256 || grid.some((n) => !bounded(n) || n === 0)) return undefined
  if (table.rows.length === 0 || table.rows.length > 10_000) return undefined
  if (table.alignment !== undefined && table.alignment !== 'left') return undefined
  const indent = table.indent_twips ?? 0
  if (!bounded(indent)) return undefined
  const gridSum = grid.reduce((a, b) => a + b, 0)
  const available = containerWidth - indent
  if (!bounded(gridSum) || gridSum === 0 || available <= 0 || gridSum > available) return undefined
  let cells = 0
  for (const row of table.rows) {
    let column = 0
    for (const cell of row.cells) {
      if (++cells > 100_000 || !Number.isSafeInteger(cell.grid_span) || cell.grid_span < 1 || column + cell.grid_span > grid.length) return undefined
      column += cell.grid_span
    }
    if (column !== grid.length) return undefined
  }
  const percentNumerator = BigInt(containerWidth) * BigInt(percent)
  const rows = table.rows.map((row) => {
    let column = 0
    return { ...row, cells: row.cells.map((cell) => { const end = column + cell.grid_span; const projected = { ...cell, width_twips: grid.slice(column, end).reduce((a, b) => a + b, 0) }; column = end; return projected }) }
  })
  const policy: NativeDocxApproximatePercentTablePolicyV1 = {
    name: DOCX_APPROXIMATE_PERCENT_TABLE_POLICY, section_id: sectionID, container_width_twips: containerWidth, percent_fiftieths: percent,
    percent_width_twips: Number(percentNumerator / 5000n), percent_width_exact: percentNumerator % 5000n === 0n,
    source_grid_widths_twips: [...grid], source_cell_widths_twips: table.rows.map((row) => row.cells.map((cell) => cell.width_twips ?? null)),
    painted_grid_widths_twips: [...grid], painted_width_twips: gridSum,
  }
  const { width_percent_fiftieths: _percent, ...rest } = table
  return {
    policy,
    table: {
      ...rest, layout: 'fixed', width_twips: gridSum, grid_widths_twips: [...grid], rows,
      alignment: table.alignment ?? 'left',
      indent_twips: indent,
      cell_margins: table.cell_margins ?? { top_twips: 0, right_twips: 115, bottom_twips: 0, left_twips: 115 },
    },
  }
}
