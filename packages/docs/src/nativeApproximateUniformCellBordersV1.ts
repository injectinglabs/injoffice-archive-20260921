/** Approximate-preview-only projection of one uniform authored cell border set
 * onto the table that states none. Strict page paint never calls this; it is a
 * declared policy of the read-only current-layout approximate preview. No
 * source bytes change. */
import type { NativeDocxTableBorderV1, NativeDocxTableV1 } from './nativeContract.js'

export const DOCX_APPROXIMATE_UNIFORM_CELL_BORDER_POLICY = 'approximate-uniform-cell-borders-projected-v1' as const

export interface NativeDocxApproximateUniformCellBorderPolicyV1 {
  name: typeof DOCX_APPROXIMATE_UNIFORM_CELL_BORDER_POLICY
  /** The one w:tcBorders value every cell states on all four edges. */
  source_border: NativeDocxTableBorderV1
  /** How many cells stated it; every cell of the table is required to. */
  source_cell_count: number
}

const EDGES = ['top', 'right', 'bottom', 'left'] as const

function sameBorder(a: NativeDocxTableBorderV1 | undefined, b: NativeDocxTableBorderV1): boolean {
  return a !== undefined && a.style === b.style && a.size_eighth_points === b.size_eighth_points && a.color_rgb === b.color_rgb
}

/** A single paintable border: exactly what the table-level painter accepts. */
function paintable(border: NativeDocxTableBorderV1 | undefined): border is NativeDocxTableBorderV1 {
  return border !== undefined && border.style === 'single' && Number.isSafeInteger(border.size_eighth_points)
    && border.size_eighth_points > 0 && border.size_eighth_points <= 768 && /^[0-9A-F]{6}$/.test(border.color_rgb ?? '')
}

/**
 * `Cell border conflict resolution is outside v1; use unambiguous table-level
 * borders` refuses every table whose cells carry w:tcBorders, because two
 * neighbouring cells can state different borders on the edge they share and
 * ECMA-376 17.4.39 leaves which one wins to the conflict-resolution rules this
 * tier does not implement. There is one shape with no conflict to resolve:
 * every cell of the table states the SAME single border on all four of its
 * edges, and the table itself states none. Every shared edge is then stated
 * twice with the same value, so the resolution is that value whichever rule
 * applies, and the unambiguous table-level border set the refusal asks for is
 * exactly that value on all six table edges.
 *
 * Verified against Word's own PDF export of
 * `tdf135943_shapeWithText_LayoutInCell0_compat15.docx`, whose two cells each
 * state `single`/`sz=2`/`000000` on all four edges against a table that states
 * none. Word draws the full grid at that value: filled 0.24 pt bars centred on
 * x = 56.76, 207.00 and 362.52 pt and on y = 240.84 and 176.76 pt, which is the
 * table box, its one interior rule and nothing else, at the 1/300 in grid that
 * export quantizes to. sz=2 is 0.25 pt, painted here as a 250 milli-point
 * stroke on those same lines.
 *
 * What stays approximate: the projection is applied without running Word's
 * conflict resolution, so it is asserted only for the one shape where no
 * conflict exists. Anything else -- a cell that omits an edge, two cells that
 * disagree, a table that already states its own borders, a vertically merged
 * cell whose continuation rows are painted by the owner -- returns undefined
 * and keeps the existing refusal on both tiers.
 */
export function projectNativeDocxApproximateUniformCellBordersV1(table: NativeDocxTableV1): { table: NativeDocxTableV1; policy: NativeDocxApproximateUniformCellBorderPolicyV1 } | undefined {
  if (table.borders !== undefined || table.rows.length === 0 || table.rows.length > 10_000) return undefined
  let border: NativeDocxTableBorderV1 | undefined
  let count = 0
  for (const row of table.rows) {
    if (row.cells.length === 0) return undefined
    for (const cell of row.cells) {
      if (++count > 100_000 || cell.vertical_merge !== 'none') return undefined
      const borders = cell.borders
      if (!borders || Object.keys(borders).sort().join(',') !== 'bottom,left,right,top') return undefined
      border ??= borders.top
      if (!paintable(border) || EDGES.some((edge) => !sameBorder(borders[edge], border!))) return undefined
    }
  }
  if (!paintable(border)) return undefined
  const projected: NativeDocxTableBorderV1 = { style: border.style, size_eighth_points: border.size_eighth_points, ...(border.color_rgb === undefined ? {} : { color_rgb: border.color_rgb }) }
  return {
    policy: { name: DOCX_APPROXIMATE_UNIFORM_CELL_BORDER_POLICY, source_border: { ...projected }, source_cell_count: count },
    table: {
      ...table,
      borders: { top: { ...projected }, right: { ...projected }, bottom: { ...projected }, left: { ...projected }, inside_horizontal: { ...projected }, inside_vertical: { ...projected } },
      rows: table.rows.map((row) => ({ ...row, cells: row.cells.map(({ borders: _borders, ...cell }) => cell) })),
    },
  }
}
