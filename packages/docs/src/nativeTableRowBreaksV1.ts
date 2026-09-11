/** Pure line-safe cuts through a natural-height row. No browser measurement. */
import type { NativeDocxQualifiedTableV1, NativeDocxTableRowGeometryV1 } from './nativeTablePagePaintV1.js'
import type { NativeDocxShapedParagraphV1 } from './nativeShapingLines.js'
import type { NativeDocxResolvedParagraphV1 } from './nativeResolvedLayout.js'

export interface NativeDocxRowParagraphV1 {
  cell_id: string
  content_x: number
  content_width: number
  paragraph: NativeDocxShapedParagraphV1
  line_y: number[]
}
export interface NativeDocxRowBreakPlanV1 {
  height: number
  paragraphs: NativeDocxRowParagraphV1[]
  events: Array<{ paragraph: number; line: number; y: number }>
  protected: Array<{ start: number; end: number }>
}

export function nativeDocxRowBreakPlanV1(table: NativeDocxQualifiedTableV1, row: NativeDocxTableRowGeometryV1, shaped: Map<string, NativeDocxShapedParagraphV1>, resolved: Map<string, NativeDocxResolvedParagraphV1>): NativeDocxRowBreakPlanV1 | undefined {
  const paragraphs: NativeDocxRowParagraphV1[] = [], protectedRanges: Array<{ start: number; end: number }> = []
  for (const [index, cell] of table.rows[row.row_ordinal]!.cells.entries()) {
    const geometry = row.cells[index]!
    let y = geometry.content_y_millipoints, previousAfter = 0
    for (const [paragraphIndex, source] of cell.cell.paragraphs.entries()) {
      const paragraph = shaped.get(source.id), properties = resolved.get(source.id)?.properties
      if (!paragraph || !properties || properties.keep_next || properties.page_break_before) return undefined
      y += paragraphIndex === 0 ? paragraph.spacing_before_millipoints : Math.max(previousAfter, paragraph.spacing_before_millipoints)
      const line_y: number[] = []
      for (const line of paragraph.lines) {
        line_y.push(y); protectedRanges.push({ start: y, end: y + line.line_height_millipoints }); y += line.line_height_millipoints
      }
      const count = paragraph.lines.length
      if (count === 0) return undefined
      if (properties.keep_lines) protectedRanges.push({ start: line_y[0]!, end: y })
      else if (properties.widow_control !== false && count > 1) {
        protectedRanges.push({ start: line_y[0]!, end: line_y[1]! + paragraph.lines[1]!.line_height_millipoints })
        protectedRanges.push({ start: line_y[count-2]!, end: y })
      }
      paragraphs.push({ cell_id: cell.cell.id, content_x: geometry.content_x_millipoints, content_width: geometry.content_width_millipoints, paragraph, line_y })
      previousAfter = paragraph.spacing_after_millipoints
    }
  }
  if (!protectedRanges.length) return undefined
  protectedRanges.sort((a,b)=>a.start-b.start || a.end-b.end)
  const merged: typeof protectedRanges = []
  for (const range of protectedRanges) {
    if (![range.start,range.end].every(Number.isSafeInteger) || range.start < 0 || range.end > row.height_millipoints || range.end <= range.start) return undefined
    const previous = merged.at(-1)
    if (previous && range.start < previous.end) previous.end = Math.max(previous.end, range.end)
    else merged.push({ ...range })
  }
  // Never emit padding-only pages: retain leading/trailing margins with ink.
  merged[0]!.start = 0; merged.at(-1)!.end = row.height_millipoints
  const events = paragraphs.flatMap((entry, paragraph) => entry.line_y.map((y, line) => ({ paragraph, line, y })))
  events.sort((a, b) => a.y - b.y || a.paragraph - b.paragraph || a.line - b.line)
  return { height: row.height_millipoints, paragraphs, events, protected: merged }
}

export function nativeDocxRowCutV1(plan: NativeDocxRowBreakPlanV1, start: number, available: number): number {
  let end = Math.min(plan.height, start + available), low = 0, high = plan.protected.length
  while (low < high) { const middle = Math.floor((low+high)/2); if (plan.protected[middle]!.start < end) low = middle+1; else high = middle }
  const interval = plan.protected[low-1]
  if (interval && end < interval.end) end = interval.start
  return Math.max(start,end)
}
