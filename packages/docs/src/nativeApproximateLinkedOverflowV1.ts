/** Approximate linked-textbox story overflow for the read-only page preview.
 *
 * Word reflows remaining story into each successor box at that box's inner
 * width. Wrapping the whole chain at the head width and then packing lines by
 * height is not that model: a later, wider (or narrower) box must re-wrap.
 * This module is disclosed approximate layout, not a Word-measured overflow. */
import type { NativeDocxParagraphV1, NativeDocxRunV1 } from './nativeContract.js'
import type { NativeDocxResolvedParagraphV1, NativeDocxResolvedRunV1 } from './nativeResolvedLayout.js'

export const EMU_PER_MILLIPOINT = 12.7

export function nativeDocxApproximateEmuToMillipointsV1(emu: number): number {
  return Math.round(emu / EMU_PER_MILLIPOINT)
}

export interface NativeDocxApproximateLinkedOverflowSlotV1 {
  /** Width used to shape remaining story for this slot (inner width, or a huge value when wrap is none). */
  wrap_width_millipoints: number
  inner_height_millipoints: number
}

export interface NativeDocxApproximateLinkedOverflowFragmentV1 {
  source_kind: string
  source_id: string
  text: string
}

export interface NativeDocxApproximateLinkedOverflowShapedLineV1 {
  line_height_millipoints: number
  fragments: readonly NativeDocxApproximateLinkedOverflowFragmentV1[]
  hard_break_after?: { source_run_id: string }
}

export interface NativeDocxApproximateLinkedOverflowShapedParagraphV1<TLine extends NativeDocxApproximateLinkedOverflowShapedLineV1 = NativeDocxApproximateLinkedOverflowShapedLineV1> {
  paragraph_id: string
  spacing_before_millipoints: number
  spacing_after_millipoints: number
  lines: readonly TLine[]
}

export interface NativeDocxApproximateLinkedOverflowStoryV1 {
  paragraphs: NativeDocxParagraphV1[]
  resolved_paragraphs: NativeDocxResolvedParagraphV1[]
  resolved_runs: NativeDocxResolvedRunV1[]
}

export interface NativeDocxApproximateLinkedOverflowPlacementV1<TLine extends NativeDocxApproximateLinkedOverflowShapedLineV1 = NativeDocxApproximateLinkedOverflowShapedLineV1> {
  slot: number
  y_millipoints: number
  paragraph_id: string
  line: TLine
}

export interface NativeDocxApproximateLinkedOverflowResultV1<TLine extends NativeDocxApproximateLinkedOverflowShapedLineV1 = NativeDocxApproximateLinkedOverflowShapedLineV1> {
  placements: NativeDocxApproximateLinkedOverflowPlacementV1<TLine>[]
  dropped_lines: number
  used_height_millipoints: number[]
}

/** Inner content frame of a text box, millipoints. Insets are left, top, right, bottom EMU. */
export function nativeDocxApproximateTextboxInnerFrameV1(
  x_millipoints: number,
  y_millipoints: number,
  width_millipoints: number,
  height_millipoints: number,
  insets_emu: readonly [number, number, number, number],
): { left: number; top: number; bottom: number; inner_width_millipoints: number; inner_height_millipoints: number } {
  const leftInset = nativeDocxApproximateEmuToMillipointsV1(insets_emu[0])
  const topInset = nativeDocxApproximateEmuToMillipointsV1(insets_emu[1])
  const rightInset = nativeDocxApproximateEmuToMillipointsV1(insets_emu[2])
  const bottomInset = nativeDocxApproximateEmuToMillipointsV1(insets_emu[3])
  return {
    left: x_millipoints + leftInset,
    top: y_millipoints + topInset,
    bottom: y_millipoints + height_millipoints - bottomInset,
    inner_width_millipoints: Math.max(1, width_millipoints - leftInset - rightInset),
    inner_height_millipoints: height_millipoints - topInset - bottomInset,
  }
}

export function consumeNativeDocxApproximateLinkedOverflowSlotV1<TLine extends NativeDocxApproximateLinkedOverflowShapedLineV1>(
  inner_height_millipoints: number,
  paragraphs: readonly NativeDocxApproximateLinkedOverflowShapedParagraphV1<TLine>[],
): {
  placements: Array<{ paragraph_index: number; line_index: number; paragraph_id: string; y_millipoints: number; line: TLine }>
  used_height_millipoints: number
  stopped: { paragraph_index: number; line_index: number } | 'complete'
} {
  const placements: Array<{ paragraph_index: number; line_index: number; paragraph_id: string; y_millipoints: number; line: TLine }> = []
  if (inner_height_millipoints <= 0) return { placements, used_height_millipoints: 0, stopped: paragraphs.length ? { paragraph_index: 0, line_index: 0 } : 'complete' }
  let cursor = 0
  for (let paragraphIndex = 0; paragraphIndex < paragraphs.length; paragraphIndex += 1) {
    const paragraph = paragraphs[paragraphIndex]!
    for (let lineIndex = 0; lineIndex < paragraph.lines.length; lineIndex += 1) {
      const line = paragraph.lines[lineIndex]!
      const extra = lineIndex === 0 ? paragraph.spacing_before_millipoints : 0
      const height = line.line_height_millipoints > 0 ? line.line_height_millipoints : 1
      if (cursor + extra + height > inner_height_millipoints) return { placements, used_height_millipoints: cursor, stopped: { paragraph_index: paragraphIndex, line_index: lineIndex } }
      cursor += extra
      placements.push({ paragraph_index: paragraphIndex, line_index: lineIndex, paragraph_id: paragraph.paragraph_id, y_millipoints: cursor, line })
      cursor += height
    }
    cursor += paragraph.spacing_after_millipoints
  }
  return { placements, used_height_millipoints: cursor, stopped: 'complete' }
}

/** Drop already-placed story. Continuation paragraphs lose first-line indent, space-before and numbering so a mid-paragraph overflow does not invent a new paragraph. */
export function remainderNativeDocxApproximateLinkedOverflowStoryV1<TLine extends NativeDocxApproximateLinkedOverflowShapedLineV1>(
  story: NativeDocxApproximateLinkedOverflowStoryV1,
  shaped: readonly NativeDocxApproximateLinkedOverflowShapedParagraphV1<TLine>[],
  stopped: { paragraph_index: number; line_index: number } | 'complete',
): NativeDocxApproximateLinkedOverflowStoryV1 {
  if (stopped === 'complete') return { paragraphs: [], resolved_paragraphs: [], resolved_runs: [] }
  const tail = shaped.slice(stopped.paragraph_index)
  if (!tail.length) return { paragraphs: [], resolved_paragraphs: [], resolved_runs: [] }
  const paragraphs: NativeDocxParagraphV1[] = []
  const resolvedParagraphs: NativeDocxResolvedParagraphV1[] = []
  const runIDs = new Set<string>()
  for (const [offset, shapedParagraph] of tail.entries()) {
    const source = story.paragraphs.find(entry => entry.id === shapedParagraph.paragraph_id)
    const resolved = story.resolved_paragraphs.find(entry => entry.paragraph_id === shapedParagraph.paragraph_id)
    if (!source) continue
    const lineIndex = offset === 0 ? stopped.line_index : 0
    const continuation = offset === 0 && lineIndex > 0
    const unplaced = shapedParagraph.lines.slice(lineIndex)
    const paragraph = continuation ? sliceParagraphFromUnplacedLines(source, unplaced) : source
    if (!paragraph) continue
    paragraphs.push(paragraph)
    if (resolved) {
      if (continuation) {
        const properties = { ...resolved.properties }
        delete properties.first_line_twips
        delete properties.spacing_before_twips
        resolvedParagraphs.push({ ...resolved, properties, numbering: undefined })
      } else resolvedParagraphs.push(resolved)
    }
    for (const run of paragraph.runs) runIDs.add(run.id)
  }
  return {
    paragraphs,
    resolved_paragraphs: resolvedParagraphs,
    resolved_runs: story.resolved_runs.filter(entry => runIDs.has(entry.run_id)),
  }
}

function sliceParagraphFromUnplacedLines<TLine extends NativeDocxApproximateLinkedOverflowShapedLineV1>(source: NativeDocxParagraphV1, unplaced: readonly TLine[]): NativeDocxParagraphV1 | undefined {
  if (!unplaced.length) return undefined
  const byID = new Map(source.runs.map(run => [run.id, run]))
  const runs: NativeDocxRunV1[] = []
  const pushRun = (run: NativeDocxRunV1, text?: string) => {
    const last = runs[runs.length - 1]
    if (text !== undefined && last && last.id === run.id && last.kind === 'text') {
      last.text = (last.text ?? '') + text
      return
    }
    if (runs.some(existing => existing.id === run.id) && run.kind !== 'text') return
    runs.push(text !== undefined && run.kind === 'text' ? { ...run, text } : { ...run })
  }
  for (const line of unplaced) {
    for (const fragment of line.fragments) {
      if (fragment.source_kind === 'list-marker') continue
      const run = byID.get(fragment.source_id)
      if (!run) continue
      if (fragment.source_kind === 'tab' || run.kind === 'control') pushRun(run)
      else if (run.kind === 'text') pushRun(run, fragment.text)
    }
    if (line.hard_break_after) {
      const run = byID.get(line.hard_break_after.source_run_id)
      if (run) pushRun(run)
    }
  }
  return { ...source, runs }
}

/** Fill linked slots in seq order, re-shaping remaining story at each slot's wrap width. */
export async function overflowNativeDocxApproximateLinkedStoryV1<TLine extends NativeDocxApproximateLinkedOverflowShapedLineV1>(
  slots: readonly NativeDocxApproximateLinkedOverflowSlotV1[],
  initial: NativeDocxApproximateLinkedOverflowStoryV1,
  shape: (wrap_width_millipoints: number, story: NativeDocxApproximateLinkedOverflowStoryV1) => Promise<readonly NativeDocxApproximateLinkedOverflowShapedParagraphV1<TLine>[]>,
): Promise<NativeDocxApproximateLinkedOverflowResultV1<TLine>> {
  const placements: NativeDocxApproximateLinkedOverflowPlacementV1<TLine>[] = []
  const used_height_millipoints = slots.map(() => 0)
  let story = initial
  let dropped_lines = 0
  for (let slot = 0; slot < slots.length; slot += 1) {
    if (!story.paragraphs.length) break
    const frame = slots[slot]!
    const shaped = await shape(Math.max(1, frame.wrap_width_millipoints), story)
    const consumed = consumeNativeDocxApproximateLinkedOverflowSlotV1(frame.inner_height_millipoints, shaped)
    if (!consumed.placements.length) {
      if (slot === slots.length - 1) dropped_lines += shaped.reduce((sum, paragraph) => sum + paragraph.lines.length, 0)
      continue
    }
    used_height_millipoints[slot] = consumed.used_height_millipoints
    for (const placement of consumed.placements) placements.push({ slot, y_millipoints: placement.y_millipoints, paragraph_id: placement.paragraph_id, line: placement.line })
    story = remainderNativeDocxApproximateLinkedOverflowStoryV1(story, shaped, consumed.stopped)
    if (slot === slots.length - 1 && consumed.stopped !== 'complete') {
      const partial = shaped[consumed.stopped.paragraph_index]
      dropped_lines += (partial ? partial.lines.length - consumed.stopped.line_index : 0) + shaped.slice(consumed.stopped.paragraph_index + 1).reduce((sum, paragraph) => sum + paragraph.lines.length, 0)
    }
  }
  if (story.paragraphs.length && slots.length === 0) dropped_lines += story.paragraphs.length
  return { placements, dropped_lines, used_height_millipoints }
}
