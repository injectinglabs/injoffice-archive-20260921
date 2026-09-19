import { describe, expect, it } from 'vitest'
import type { NativeDocxSectionV1 } from './nativeContract.js'
import type { NativeDocxShapedLinesV1 } from './nativeShapingLines.js'
import type { NativeDocxPaintPageV1 } from './nativePagePaintV1.js'
import {
  paintNativeDocxApproximateColumnSeparatorsV1,
  DOCX_APPROXIMATE_COLUMN_SEPARATOR_TABLE_ID,
  DOCX_APPROXIMATE_COLUMN_SEPARATOR_WIDTH_MILLIPOINTS,
} from './nativeApproximateColumnSeparatorsV1.js'

// multi-column-separator-with-line.docx, value for value: pgSz 12240x15840,
// pgMar left/right 1008 top/bottom 1296, two equal columns with w:space 720.
const COLUMN_WIDTH = 4752 * 50
const BODY_X = 1008 * 50
const BODY_TOP = 1296 * 50
const GAP = 720 * 50
const LINE_HEIGHT = 15_440
const SPACE_AFTER = 10_000

function section(overrides: Partial<NativeDocxSectionV1['page']> = {}): NativeDocxSectionV1 {
  return {
    id: 'section:1',
    page: {
      width_twips: 12240, height_twips: 15840, orientation: 'portrait',
      margins: { top_twips: 1296, right_twips: 1008, bottom_twips: 1296, left_twips: 1008, header_twips: 720, footer_twips: 720, gutter_twips: 0 },
      columns: 2, column_spacing_twips: 720, column_layout: 'equal-width',
      column_definitions: [{ id: 'column:1', ordinal: 0 }, { id: 'column:2', ordinal: 1 }],
      column_separator: true,
      ...overrides,
    },
  } as unknown as NativeDocxSectionV1
}

function page(lines: Array<{ paragraph_id: string; column_id: string; x: number; y: number }>): NativeDocxPaintPageV1 {
  return {
    id: 'page:1', ordinal: 0, section_id: 'section:1', section_ids: ['section:1'], kind: 'content',
    width_millipoints: 12240 * 50, height_millipoints: 15840 * 50,
    body_box: { x_millipoints: BODY_X, y_millipoints: BODY_TOP, width_millipoints: 10224 * 50, height_millipoints: 13248 * 50 },
    columns: [
      { id: 'column:1', section_id: 'section:1', ordinal: 0, x_millipoints: BODY_X, y_millipoints: BODY_TOP, width_millipoints: COLUMN_WIDTH, height_millipoints: 13248 * 50 },
      { id: 'column:2', section_id: 'section:1', ordinal: 1, x_millipoints: BODY_X + COLUMN_WIDTH + GAP, y_millipoints: BODY_TOP, width_millipoints: COLUMN_WIDTH, height_millipoints: 13248 * 50 },
    ],
    background_rgb: 'FFFFFF', clip_box: { x_millipoints: 0, y_millipoints: 0, width_millipoints: 12240 * 50, height_millipoints: 15840 * 50 },
    lines: lines.map((line, index) => ({
      placed_line_id: `placed:${index}`, line_id: `line:${index}`, paragraph_id: line.paragraph_id, region: 'body',
      section_id: 'section:1', column_id: line.column_id, column_ordinal: line.column_id === 'column:1' ? 0 : 1,
      source_line_ordinal: 0, x_millipoints: line.x, y_millipoints: line.y, width_millipoints: 20_000,
      height_millipoints: LINE_HEIGHT, baseline_y_millipoints: line.y + 12_000, command_ids: [],
    })),
    glyph_outlines: [], commands: [],
  } as unknown as NativeDocxPaintPageV1
}

function shaped(spacing = SPACE_AFTER): NativeDocxShapedLinesV1 {
  return { paragraphs: [
    { paragraph_id: 'paragraph:1', spacing_after_millipoints: spacing },
    { paragraph_id: 'paragraph:2', spacing_after_millipoints: spacing },
  ] } as unknown as NativeDocxShapedLinesV1
}

const BALANCED = [
  { paragraph_id: 'paragraph:1', column_id: 'column:1', x: BODY_X, y: BODY_TOP },
  { paragraph_id: 'paragraph:2', column_id: 'column:2', x: BODY_X + COLUMN_WIDTH + GAP, y: BODY_TOP },
]

describe('approximate column separators', () => {
  it('paints Word-measured geometry: a 0.96 pt bar centred on the gap, from the column top to the content bottom', () => {
    const target = page(BALANCED)
    const result = paintNativeDocxApproximateColumnSeparatorsV1([target], [section()], shaped())
    expect(result.painted).toEqual(['section:1'])
    expect(target.commands).toHaveLength(1)
    const command = target.commands[0]!
    expect(command.kind).toBe('stroke_table_border')
    if (command.kind !== 'stroke_table_border') throw new Error('expected a stroke')
    expect(command.table_id).toBe(DOCX_APPROXIMATE_COLUMN_SEPARATOR_TABLE_ID)
    expect(command.width_millipoints).toBe(DOCX_APPROXIMATE_COLUMN_SEPARATOR_WIDTH_MILLIPOINTS)
    expect(command.stroke_rgb).toBe('000000')
    // Word: a bar centred on x = 306 pt between column origins 50.4 pt and
    // 324.0 pt. The command states the leading edge, so 306 pt - 0.48 pt.
    expect(command.x1_millipoints).toBe(306_000 - 480)
    expect(command.x2_millipoints).toBe(command.x1_millipoints)
    // Word: y = 64.8 pt (the 1296-twip top margin) to 90.24 pt (one 15.44 pt
    // line plus that paragraph's 10 pt space-after).
    expect(command.y1_millipoints).toBe(64_800)
    expect(command.y2_millipoints).toBe(64_800 + LINE_HEIGHT + SPACE_AFTER)
  })

  it('paints nothing when the section does not declare the rule', () => {
    const target = page(BALANCED)
    const result = paintNativeDocxApproximateColumnSeparatorsV1([target], [section({ column_separator: undefined })], shaped())
    expect(result.painted).toEqual([])
    expect(target.commands).toEqual([])
  })

  it('paints nothing for a single-column section', () => {
    const target = page(BALANCED)
    target.columns = [target.columns[0]!]
    const result = paintNativeDocxApproximateColumnSeparatorsV1([target], [section({ columns: 1 })], shaped())
    expect(result.painted).toEqual([])
    expect(target.commands).toEqual([])
  })

  it('paints nothing on a page that placed no body line of the section', () => {
    const target = page([])
    const result = paintNativeDocxApproximateColumnSeparatorsV1([target], [section()], shaped())
    expect(result.painted).toEqual([])
    expect(target.commands).toEqual([])
  })

  it('paints one bar per gap and keeps them inside every gap', () => {
    const target = page(BALANCED)
    target.columns = [
      { id: 'column:1', section_id: 'section:1', ordinal: 0, x_millipoints: BODY_X, y_millipoints: BODY_TOP, width_millipoints: 100_000, height_millipoints: 600_000 },
      { id: 'column:2', section_id: 'section:1', ordinal: 1, x_millipoints: BODY_X + 136_000, y_millipoints: BODY_TOP, width_millipoints: 100_000, height_millipoints: 600_000 },
      { id: 'column:3', section_id: 'section:1', ordinal: 2, x_millipoints: BODY_X + 272_000, y_millipoints: BODY_TOP, width_millipoints: 100_000, height_millipoints: 600_000 },
    ] as NativeDocxPaintPageV1['columns']
    const result = paintNativeDocxApproximateColumnSeparatorsV1([target], [section({ columns: 3 })], shaped())
    expect(result.painted).toEqual(['section:1'])
    expect(target.commands).toHaveLength(2)
    for (const command of target.commands) {
      if (command.kind !== 'stroke_table_border') throw new Error('expected a stroke')
      expect(command.x2_millipoints - command.x1_millipoints).toBe(0)
      expect(command.x1_millipoints % 1).toBe(0)
    }
  })

  it('leaves the deepest column short when only a shallower paragraph carries space-after', () => {
    const target = page([
      { paragraph_id: 'paragraph:1', column_id: 'column:1', x: BODY_X, y: BODY_TOP },
      { paragraph_id: 'paragraph:2', column_id: 'column:2', x: BODY_X + COLUMN_WIDTH + GAP, y: BODY_TOP + 100_000 },
    ])
    const spacings = { paragraphs: [
      { paragraph_id: 'paragraph:1', spacing_after_millipoints: 500_000 },
      { paragraph_id: 'paragraph:2', spacing_after_millipoints: 0 },
    ] } as unknown as NativeDocxShapedLinesV1
    const result = paintNativeDocxApproximateColumnSeparatorsV1([target], [section()], spacings)
    expect(result.painted).toEqual(['section:1'])
    const command = target.commands[0]!
    if (command.kind !== 'stroke_table_border') throw new Error('expected a stroke')
    expect(command.y2_millipoints).toBe(BODY_TOP + 100_000 + LINE_HEIGHT)
  })
})
