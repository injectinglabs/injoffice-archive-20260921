/** Bounded content-based table sizing. No browser metrics or source rewriting. */
import type { NativeDocxTableV1 } from './nativeContract.js'
import type { NativeDocxResolvedLayoutInputV1 } from './nativeResolvedLayout.js'
import type { NativeDocxShapedLinesV1 } from './nativeShapingLines.js'

export interface NativeDocxTableAutofitPolicyV1 {
  name: 'shaped-content-minmax-v1' | 'source-preferred-nonconflicting-v1'
  section_id: string
  container_width_twips: number
  source_grid_widths_twips: number[]
  source_cell_widths_twips: Array<Array<number | null>>
  preferred_width_twips: number | null
  minimum_widths_twips: number[]
  maximum_widths_twips: number[]
}

const LIMIT = 20_000_000
const bounded = (n: unknown): n is number => typeof n === 'number' && Number.isSafeInteger(n) && n >= 0 && n <= LIMIT

/** Re-derived both before final shaping and from final, source-bound clusters.
 * Soft wraps do not change intrinsic widths. Hard breaks end max-content lines.
 * Only U+0020 is a word separator in this conservative first policy.
 */
export function resolveNativeDocxTableAutofitV1(table: NativeDocxTableV1, containerWidth: number, sectionID: string, resolved: NativeDocxResolvedLayoutInputV1, shaped?: NativeDocxShapedLinesV1, indexes?: { paragraphs: ReadonlyMap<string, NativeDocxShapedLinesV1['paragraphs'][number]>; properties: ReadonlyMap<string, NativeDocxResolvedLayoutInputV1['paragraphs'][number]> }): { table: NativeDocxTableV1; policy: NativeDocxTableAutofitPolicyV1 } | undefined {
  const grid = table.grid_widths_twips, margins = table.cell_margins
  if (!shaped || !grid?.length || grid.length > 64 || grid.some(n => !bounded(n) || n === 0) || !margins || !bounded(containerWidth) || !bounded(table.indent_twips) || table.width_percent_fiftieths !== undefined || table.rows.length > 1_000 || table.rows.length === 0) return undefined
  if (table.width_twips !== undefined && (!bounded(table.width_twips) || table.width_twips === 0)) return undefined
  if (![margins.left_twips, margins.right_twips].every(bounded)) return undefined
  const padding = margins.left_twips + margins.right_twips
  const available = containerWidth - table.indent_twips
  const paragraphs = indexes?.paragraphs ?? new Map(shaped.paragraphs.map(p => [p.paragraph_id, p]))
  const properties = indexes?.properties ?? new Map(resolved.paragraphs.map(p => [p.paragraph_id, p]))
  const min = grid.map(() => padding + 1), max = [...min]
  let fragments = 0, paragraphCount = 0
  for (const row of table.rows) {
    if (row.cells.length !== grid.length) return undefined
    for (const [column, cell] of row.cells.entries()) {
      if (cell.grid_span !== 1 || cell.vertical_merge !== 'none' || cell.width_twips !== undefined && !bounded(cell.width_twips)) return undefined
      for (const source of cell.paragraphs) {
        if (++paragraphCount > 10_000 || source.runs.some(run => run.kind !== 'text' && !(run.kind === 'control' && run.control === 'line-break'))) return undefined
        const paragraph = paragraphs.get(source.id), p = properties.get(source.id)
        if (!paragraph || !p || p.numbering || paragraph.list_marker || paragraph.direction !== 'ltr' || !['left', 'start'].includes(paragraph.alignment) || paragraph.indent_start_millipoints || paragraph.indent_end_millipoints || paragraph.first_line_delta_millipoints) return undefined
        let word = 0, lineWidth = 0, minimum = 0, maximum = 0
        for (const line of paragraph.lines) {
          if (line.fragments.length > 100_000 - fragments) return undefined
          for (const fragment of [...line.fragments].sort((a,b) => a.logical_order - b.logical_order)) {
            if (++fragments > 100_000 || fragment.source_kind !== 'run' || fragment.justification_expansion_millipoints || !Number.isSafeInteger(fragment.advance_inline_millipoints) || fragment.advance_inline_millipoints < 0) return undefined
            // Tabs, NBSP, discretionary breaks, and bidi controls do not share
            // this policy's simple whitespace opportunity model.
            if (/[\t\r\n\u00a0\u00ad\u2000-\u206f]/u.test(fragment.text) || fragment.text.includes(' ') && !/^ +$/.test(fragment.text)) return undefined
            word += fragment.advance_inline_millipoints
            lineWidth += fragment.advance_inline_millipoints
            if (word > LIMIT * 50 || lineWidth > LIMIT * 50) return undefined
            if (/^ +$/.test(fragment.text)) { minimum = Math.max(minimum, word); word = 0 }
          }
          if (line.hard_break_after) { minimum = Math.max(minimum, word); maximum = Math.max(maximum, lineWidth); word = 0; lineWidth = 0 }
        }
        minimum = Math.max(minimum, word); maximum = Math.max(maximum, lineWidth)
        min[column] = Math.max(min[column]!, Math.ceil(minimum / 50) + padding)
        max[column] = Math.max(max[column]!, Math.ceil(maximum / 50) + padding, min[column]!)
      }
    }
  }
  const minimum = min.reduce((a,b) => a+b,0), maximum = max.reduce((a,b) => a+b,0)
  if (minimum > available || maximum > LIMIT || available <= 0) return undefined
  // Auto width is not permission to discard consistent authored preferences.
  // Qualify only the non-conflicting case: every cell repeats its grid width,
  // all unwrapped content fits that preference, and the complete grid fits the
  // section. This is not a general Word autofit algorithm or a fixed-grid
  // override: explicit table widths and conflicting/wrapping preferences keep
  // the existing content policy below. Final source-bound shaping replays this
  // same decision, and the named policy enters the table projection hash.
  const preferredGridWidth = grid.reduce((a,b) => a+b,0)
  const preservePreferences = table.width_twips === undefined
    && preferredGridWidth <= available
    && grid.every((width,index) => max[index]! <= width)
    && table.rows.every(row => row.cells.every((cell,index) => cell.width_twips === grid[index]))
  if (preservePreferences) {
    const policy: NativeDocxTableAutofitPolicyV1 = { name:'source-preferred-nonconflicting-v1', section_id:sectionID, container_width_twips:containerWidth, source_grid_widths_twips:[...grid], source_cell_widths_twips:table.rows.map(row=>row.cells.map(cell=>cell.width_twips??null)), preferred_width_twips:null, minimum_widths_twips:min, maximum_widths_twips:max }
    return { policy, table:{ ...table, layout:'fixed', width_twips:preferredGridWidth, grid_widths_twips:[...grid], rows:table.rows.map(row=>({ ...row,cells:row.cells.map((cell,index)=>({...cell,width_twips:grid[index]!})) })) } }
  }
  // Authored table width is a preferred target, clamped to intrinsic min/max
  // and the source section. Cell/grid preferences are preserved in provenance,
  // not treated as immutable column widths (that would be fixed-grid sizing).
  const target = Math.min(available, maximum, Math.max(minimum, table.width_twips ?? maximum))
  const room = max.map((n,i) => n-min[i]!), totalRoom = maximum-minimum, extra = target-minimum
  const widths = [...min]
  if (totalRoom > 0) {
    const remainders = room.map((n,index) => { const numerator = BigInt(n)*BigInt(extra); widths[index]! += Number(numerator/BigInt(totalRoom)); return { index, remainder: numerator%BigInt(totalRoom) } })
    remainders.sort((a,b) => a.remainder === b.remainder ? a.index-b.index : a.remainder > b.remainder ? -1 : 1)
    const left = target-widths.reduce((a,b) => a+b,0)
    for (let i=0;i<left;i+=1) widths[remainders[i]!.index]! += 1
  }
  const policy: NativeDocxTableAutofitPolicyV1 = { name:'shaped-content-minmax-v1', section_id:sectionID, container_width_twips:containerWidth, source_grid_widths_twips:[...grid], source_cell_widths_twips:table.rows.map(row=>row.cells.map(cell=>cell.width_twips??null)), preferred_width_twips:table.width_twips??null, minimum_widths_twips:min, maximum_widths_twips:max }
  return { policy, table:{ ...table, layout:'fixed', width_twips:target, grid_widths_twips:widths, rows:table.rows.map(row=>({ ...row,cells:row.cells.map((cell,index)=>({...cell,width_twips:widths[index]!})) })) } }
}
