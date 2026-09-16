import { describe, expect, it } from 'vitest'
import type { NativeDocxParagraphV1 } from './nativeContract.js'
import type { NativeDocxResolvedParagraphV1, NativeDocxResolvedRunV1 } from './nativeResolvedLayout.js'
import {
  consumeNativeDocxApproximateLinkedOverflowSlotV1,
  nativeDocxApproximateTextboxInnerFrameV1,
  overflowNativeDocxApproximateLinkedStoryV1,
  remainderNativeDocxApproximateLinkedOverflowStoryV1,
  type NativeDocxApproximateLinkedOverflowShapedLineV1,
  type NativeDocxApproximateLinkedOverflowShapedParagraphV1,
  type NativeDocxApproximateLinkedOverflowStoryV1,
} from './nativeApproximateLinkedOverflowV1.js'

const HASH = `sha256:${'a'.repeat(64)}`
const POLICY = { mode: 'read-only' as const, allowed_operations: [] as const, refusal: { code: 'NATIVE_READ_ONLY', message: 'Fixture is immutable.', preservation: 'refuse-mutation' as const } }
const LINE = 8_000
const UNIT = 1_000

function anchor(path: string, start: number, end: number) {
  return { part_name: 'word/document.xml', path, start_byte: start, end_byte: end, xml_sha256: HASH }
}

function story(text: string, id = 'paragraph:1'): NativeDocxApproximateLinkedOverflowStoryV1 {
  const paragraph: NativeDocxParagraphV1 = {
    id, anchor: anchor('/w:document[1]/w:body[1]/w:p[1]', 1, 40), edit_policy: POLICY, properties: {},
    runs: [{ kind: 'text', id: `${id}:r0`, anchor: anchor('/w:document[1]/w:body[1]/w:p[1]/w:r[1]', 2, 30), text }],
  }
  const resolved_paragraphs: NativeDocxResolvedParagraphV1[] = [{ paragraph_id: id, applied_styles: [], properties: {}, paragraph_mark_properties: {} }]
  const resolved_runs: NativeDocxResolvedRunV1[] = [{ run_id: `${id}:r0`, paragraph_id: id, applied_paragraph_styles: [], applied_character_styles: [], properties: {} }]
  return { paragraphs: [paragraph], resolved_paragraphs, resolved_runs }
}

interface FakeLine extends NativeDocxApproximateLinkedOverflowShapedLineV1 { wrap_width_millipoints: number }

function tokens(text: string): string[] {
  return text.split(/( )/g).filter(token => token.length > 0)
}

function fakeShape(width: number, remaining: NativeDocxApproximateLinkedOverflowStoryV1): NativeDocxApproximateLinkedOverflowShapedParagraphV1<FakeLine>[] {
  return remaining.paragraphs.map(paragraph => {
    const text = paragraph.runs.map(run => run.kind === 'text' ? run.text ?? '' : run.kind === 'control' && run.control === 'tab' ? '\t' : '').join('')
    const sourceID = paragraph.runs[0]?.id ?? paragraph.id
    const parts = tokens(text)
    const lines: FakeLine[] = []
    let offset = 0
    if (!parts.length) lines.push({ line_height_millipoints: LINE, fragments: [], wrap_width_millipoints: width })
    while (offset < parts.length) {
      let used = 0
      const start = offset
      while (offset < parts.length) {
        const size = parts[offset]!.length * UNIT
        if (used > 0 && used + size > width) break
        used += size
        offset += 1
      }
      if (offset === start) offset += 1
      lines.push({
        line_height_millipoints: LINE,
        fragments: parts.slice(start, offset).map(token => ({ source_kind: 'run', source_id: sourceID, text: token })),
        wrap_width_millipoints: width,
      })
    }
    return { paragraph_id: paragraph.id, spacing_before_millipoints: 0, spacing_after_millipoints: 0, lines }
  })
}

function placedText(placements: Array<{ line: FakeLine }>): string {
  return placements.map(placement => placement.line.fragments.map(fragment => fragment.text).join('')).join('')
}

describe('approximate linked textbox overflow', () => {
  it('subtracts per-box insets from wrap width and height', () => {
    // Default DrawingML insets: 0.1in left/right, 0.05in top/bottom.
    const frame = nativeDocxApproximateTextboxInnerFrameV1(10_000, 20_000, 100_000, 50_000, [91_440, 45_720, 91_440, 45_720])
    expect(frame).toEqual({ left: 17_200, top: 23_600, bottom: 66_400, inner_width_millipoints: 85_600, inner_height_millipoints: 42_800 })
  })

  it('fills one slot by height and leaves the rest of the paragraph as remainder', () => {
    const remaining = story('aaaa bbbb cccc dddd')
    remaining.resolved_paragraphs[0]!.properties.first_line_twips = 200
    remaining.resolved_paragraphs[0]!.properties.spacing_before_twips = 120
    const shaped = fakeShape(5_000, remaining)
    expect(shaped[0]!.lines).toHaveLength(4)
    const consumed = consumeNativeDocxApproximateLinkedOverflowSlotV1(17_000, shaped)
    expect(consumed.placements).toHaveLength(2)
    expect(consumed.stopped).toEqual({ paragraph_index: 0, line_index: 2 })
    const rest = remainderNativeDocxApproximateLinkedOverflowStoryV1(remaining, shaped, consumed.stopped)
    expect(rest.paragraphs[0]!.runs[0]!.text).toBe('cccc dddd')
    expect(rest.resolved_paragraphs[0]!.properties.first_line_twips).toBeUndefined()
    expect(rest.resolved_paragraphs[0]!.properties.spacing_before_twips).toBeUndefined()
    expect(rest.resolved_paragraphs[0]!.numbering).toBeUndefined()
  })

  it('re-wraps remaining story at the successor box inner width, not the head width', async () => {
    const widths: number[] = []
    const result = await overflowNativeDocxApproximateLinkedStoryV1<FakeLine>(
      [
        { wrap_width_millipoints: 5_000, inner_height_millipoints: 17_000 },
        { wrap_width_millipoints: 25_000, inner_height_millipoints: 50_000 },
      ],
      story('aaaa bbbb cccc dddd eeee ffff'),
      async (width, remaining) => { widths.push(width); return fakeShape(width, remaining) },
    )
    expect(widths).toEqual([5_000, 25_000])
    const first = result.placements.filter(placement => placement.slot === 0)
    const second = result.placements.filter(placement => placement.slot === 1)
    expect(first.map(placement => placement.slot)).toEqual([0, 0])
    expect(placedText(first)).toBe('aaaa bbbb ')
    // Head-width packing would put four short lines in the second rect; the wider slot must reflow them.
    expect(second).toHaveLength(1)
    expect(second[0]!.line.wrap_width_millipoints).toBe(25_000)
    expect(placedText(second)).toBe('cccc dddd eeee ffff')
    expect(result.dropped_lines).toBe(0)
    expect(result.placements.map(placement => placement.slot)).toEqual([0, 0, 1])
  })

  it('re-wraps remaining story when the successor box is narrower than the head', async () => {
    const widths: number[] = []
    const result = await overflowNativeDocxApproximateLinkedStoryV1<FakeLine>(
      [
        { wrap_width_millipoints: 10_000, inner_height_millipoints: 9_000 },
        { wrap_width_millipoints: 5_000, inner_height_millipoints: 50_000 },
      ],
      story('aaaa bbbb cccc dddd eeee ffff'),
      async (width, remaining) => { widths.push(width); return fakeShape(width, remaining) },
    )
    expect(widths).toEqual([10_000, 5_000])
    const first = result.placements.filter(placement => placement.slot === 0)
    const second = result.placements.filter(placement => placement.slot === 1)
    expect(first).toHaveLength(1)
    expect(first[0]!.line.wrap_width_millipoints).toBe(10_000)
    expect(second.length).toBeGreaterThan(1)
    expect(second.every(placement => placement.line.wrap_width_millipoints === 5_000)).toBe(true)
    expect(placedText(first) + placedText(second)).toBe('aaaa bbbb cccc dddd eeee ffff')
  })

  it('drops lines that do not fit the last slot and keeps seq order', async () => {
    const result = await overflowNativeDocxApproximateLinkedStoryV1<FakeLine>(
      [
        { wrap_width_millipoints: 5_000, inner_height_millipoints: 9_000 },
        { wrap_width_millipoints: 5_000, inner_height_millipoints: 9_000 },
      ],
      story('aaaa bbbb cccc dddd'),
      async (width, remaining) => fakeShape(width, remaining),
    )
    expect(result.placements.map(placement => placement.slot)).toEqual([0, 1])
    expect(placedText(result.placements)).toBe('aaaa bbbb ')
    expect(result.dropped_lines).toBe(2)
  })
})
