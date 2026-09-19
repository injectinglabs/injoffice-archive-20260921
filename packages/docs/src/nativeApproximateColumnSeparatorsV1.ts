/**
 * ECMA-376 17.6.4 `w:cols/@w:sep`: a vertical rule drawn in every inter-column
 * gap of a multi-column section.
 *
 * The rule divides no box and moves no column — the column origins a section
 * with `w:sep` paints at are value-for-value the ones this model already
 * derives from `w:pgSz`, `w:pgMar` and `w:cols` — so painting it adds ink and
 * changes nothing else about the page.
 *
 * Geometry, measured on Word's own export of
 * `office-hard-v2/pages/multi-column-separator-with-line`:
 *
 *   * a filled 0.96 pt wide bar centred on x = 306 pt, the exact centre of the
 *     gap between column origins 50.4 pt and 324.0 pt (widths 118.8 pt, gap
 *     36 pt), so the bar is centred in the gap, not aligned to either column;
 *   * it starts at the top of the column box — y = 64.8 pt, the section's
 *     1296-twip top margin — and not at the first baseline;
 *   * it ends at 90.24 pt, which is the one line the balanced column holds
 *     (15.44 pt) plus that paragraph's 10 pt `w:spacing/@w:after`. Word ends
 *     the rule at the bottom of the content the column occupies, and a
 *     paragraph occupies its lines plus its space-after; the rule is not run
 *     to the foot of the body box, which is 89 pt further down.
 *
 * What this painter will not state it refuses to paint, and the section keeps
 * its `COLUMN_SEPARATOR_UNSUPPORTED` refusal: a column whose deepest line this
 * page did not place, a gap that is not positive, or a bar that would leave
 * the page. The strict tier refuses the attribute either way.
 */

import type { NativeDocxSectionV1 } from './nativeContract.js'
import type { NativeDocxShapedLinesV1 } from './nativeShapingLines.js'
import type { NativeDocxPaintPageV1, NativeDocxPagePaintCommandV1, NativeDocxStrokeTableBorderCommandV1 } from './nativePagePaintV1.js'

/** Synthetic table identity these strokes are bound to; it names no source table. */
export const DOCX_APPROXIMATE_COLUMN_SEPARATOR_TABLE_ID = 'approximate:column-separator'

/** ECMA-376 states no width for the rule; Word paints 0.96 pt. */
export const DOCX_APPROXIMATE_COLUMN_SEPARATOR_WIDTH_MILLIPOINTS = 960

export const DOCX_APPROXIMATE_COLUMN_SEPARATOR_WARNING =
  'Approximate read-only preview: a w:cols w:sep rule is painted as a 0.96 pt bar centred in each inter-column gap, from the top of the column box to the bottom of the content it holds (approximate-column-separator-v1).'

export interface NativeDocxApproximateColumnSeparatorResultV1 {
  /** Section ids whose rule was painted on at least one page. */
  painted: string[]
  reasons: string[]
}

interface Bar {
  pageID: string
  sectionID: string
  leftColumnID: string
  /** The bar's leading (left) edge, not its axis. */
  x: number
  top: number
  bottom: number
}

/**
 * Append the rule to every page that carries a section declaring it.
 *
 * Strokes are appended to `page.commands` after the existing ones. The wire's
 * command order already places `stroke_table_border` last, and this bar is a
 * decoration under nothing: it never overlaps a column's own text box.
 */
export function paintNativeDocxApproximateColumnSeparatorsV1(
  pages: readonly NativeDocxPaintPageV1[],
  sections: readonly NativeDocxSectionV1[],
  shaped: NativeDocxShapedLinesV1,
): NativeDocxApproximateColumnSeparatorResultV1 {
  const declaring = new Set(sections.filter((section) => section.page.column_separator === true && section.page.columns > 1).map((section) => section.id))
  if (declaring.size === 0) return { painted: [], reasons: [] }
  const spacingAfter = new Map<string, number>()
  for (const paragraph of shaped.paragraphs) spacingAfter.set(paragraph.paragraph_id, paragraph.spacing_after_millipoints)
  const painted = new Set<string>()
  for (const page of pages) {
    const bars: Bar[] = []
    for (const sectionID of declaring) {
      const columns = page.columns.filter((column) => column.section_id === sectionID).sort((left, right) => left.ordinal - right.ordinal)
      if (columns.length < 2) continue
      const bottom = sectionContentBottom(page, sectionID, spacingAfter)
      if (bottom === undefined) continue
      for (let index = 0; index + 1 < columns.length; index += 1) {
        const left = columns[index]!
        const right = columns[index + 1]!
        const gapStart = left.x_millipoints + left.width_millipoints
        const gapEnd = right.x_millipoints
        if (!(gapEnd > gapStart)) continue
        // Word centres the bar in the gap. An odd gap has no whole-millipoint
        // centre; nothing states which side Word would round to, so it is not
        // painted rather than painted half a thousandth of a point off true.
        if ((gapStart + gapEnd) % 2 !== 0) continue
        // A vertical stroke command states its LEADING edge, the way a
        // horizontal one states its top, so the centred bar is emitted at
        // centre - width/2 and not at the centre itself.
        const x = (gapStart + gapEnd) / 2 - DOCX_APPROXIMATE_COLUMN_SEPARATOR_WIDTH_MILLIPOINTS / 2
        const top = left.y_millipoints
        if (!(bottom > top)) continue
        if (x < gapStart || x + DOCX_APPROXIMATE_COLUMN_SEPARATOR_WIDTH_MILLIPOINTS > gapEnd || bottom > page.height_millipoints) continue
        bars.push({ pageID: page.id, sectionID, leftColumnID: left.id, x, top, bottom })
      }
    }
    if (bars.length === 0) continue
    const commands: NativeDocxPagePaintCommandV1[] = bars.map((bar) => ({
      kind: 'stroke_table_border',
      id: `paint:column-separator:${bar.pageID}:${bar.leftColumnID}`,
      table_id: DOCX_APPROXIMATE_COLUMN_SEPARATOR_TABLE_ID,
      row_id: bar.sectionID,
      cell_id: bar.leftColumnID,
      edge: 'right',
      x1_millipoints: bar.x,
      y1_millipoints: bar.top,
      x2_millipoints: bar.x,
      y2_millipoints: bar.bottom,
      width_millipoints: DOCX_APPROXIMATE_COLUMN_SEPARATOR_WIDTH_MILLIPOINTS,
      stroke_rgb: '000000',
    } satisfies NativeDocxStrokeTableBorderCommandV1))
    page.commands = [...page.commands, ...commands]
    for (const bar of bars) painted.add(bar.sectionID)
  }
  const result: NativeDocxApproximateColumnSeparatorResultV1 = { painted: [...painted], reasons: [] }
  if (result.painted.length > 0) result.reasons.push(DOCX_APPROXIMATE_COLUMN_SEPARATOR_WARNING)
  return result
}

/**
 * The bottom of the content the section occupies on this page: the deepest
 * body line it placed, plus that line's paragraph space-after.
 *
 * A page that placed no body line for the section states no extent and paints
 * no rule.
 */
function sectionContentBottom(page: NativeDocxPaintPageV1, sectionID: string, spacingAfter: ReadonlyMap<string, number>): number | undefined {
  let deepest: NativeDocxPaintPageV1['lines'][number] | undefined
  for (const line of page.lines) {
    if (line.region !== 'body' || line.section_id !== sectionID) continue
    if (deepest === undefined || line.y_millipoints + line.height_millipoints > deepest.y_millipoints + deepest.height_millipoints) deepest = line
  }
  if (deepest === undefined) return undefined
  // A paragraph above the deepest one spends its space-after as the gap to the
  // paragraph below it, so only the deepest line's own paragraph adds depth.
  const after = spacingAfter.get(deepest.paragraph_id)
  if (after === undefined) return undefined
  return deepest.y_millipoints + deepest.height_millipoints + after
}
