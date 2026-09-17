import { describe, expect, it } from 'vitest'
import type { NativeDocxResolvedNumberingV1 } from './nativeResolvedLayout.js'
import { decodeNativeDocxResolvedLayout, nativeDocxResolvedNumberingDefinitionSha256V1 } from './nativeResolvedLayout.js'
import { nativeDocxListSuffixTabTargetV1, positionNativeDocxListMarkerV1 } from './nativeNumberingV1.js'

function marker(alignment: NativeDocxResolvedNumberingV1['alignment']): NativeDocxResolvedNumberingV1 {
  return {
    marker_id: 'marker:1', definition_sha256: `sha256:${'a'.repeat(64)}`, num_id: '1', abstract_num_id: '1', level: 0,
    start: 1, format: 'decimal', text: '%1.', suffix: 'tab', alignment, never_restart: true,
    counter_value: 1, counter_values: [{ level: 0, value: 1, format: 'decimal' }], resolved_text: '1.',
    label_start_twips: 100, label_end_twips: 300, text_start_twips: 300, marker_properties: {},
  }
}

describe('native DOCX marker geometry', () => {
  it('matches the versioned Go canonical definition vector for HTML, separators, astral, and non-ASCII text', () => {
    const value: NativeDocxResolvedNumberingV1 = {
      marker_id: 'marker:vector', definition_sha256: `sha256:${'0'.repeat(64)}`,
      num_id: '7', abstract_num_id: '3', level: 2, level_style_id: 'List<&>😀',
      start: 4, format: 'decimal', text: '<>&\u2028\u2029😀é%3', suffix: 'space', alignment: 'center', restart_after_level: 1, never_restart: false,
      counter_value: 4, counter_values: [], resolved_text: '4', numbering_tab_twips: 840,
      label_start_twips: 360, label_end_twips: 1080, text_start_twips: 1080, marker_properties: {},
    }
    expect(nativeDocxResolvedNumberingDefinitionSha256V1(value, `sha256:${'a'.repeat(64)}`)).toBe('sha256:654cfe4ae66ee21d8b0e330ec13d58eef947f7fd76378515500cdd5a6489a31d')
  })

  it('positions start, end, and centered labels using shaped width only', () => {
    expect(positionNativeDocxListMarkerV1(marker('start'), 'ltr', 2_000)?.marker_start_millipoints).toBe(5_000)
    expect(positionNativeDocxListMarkerV1(marker('end'), 'ltr', 2_000)?.marker_start_millipoints).toBe(3_000)
    expect(positionNativeDocxListMarkerV1(marker('center'), 'ltr', 2_000)?.marker_start_millipoints).toBe(4_000)
    expect(positionNativeDocxListMarkerV1(marker('start'), 'rtl', 2_000)?.marker_start_millipoints).toBe(3_000)
    expect(positionNativeDocxListMarkerV1(marker('end'), 'rtl', 2_000)?.marker_start_millipoints).toBe(5_000)
    expect(positionNativeDocxListMarkerV1(marker('right'), 'ltr', 2_000)?.marker_start_millipoints).toBe(3_000)
    expect(positionNativeDocxListMarkerV1(marker('left'), 'rtl', 2_000)?.marker_start_millipoints).toBe(5_000)
    expect(positionNativeDocxListMarkerV1(marker('center'), 'ltr', 2_001)?.marker_start_millipoints).toBe(3_999)
  })

  it('refuses overflow and resolves hanging/default suffix tab targets exactly', () => {
    expect(positionNativeDocxListMarkerV1(marker('center'), 'ltr', 10_001)).toBeUndefined()
    expect(nativeDocxListSuffixTabTargetV1(7_000, 15_000, 4_000)).toBe(15_000)
    expect(nativeDocxListSuffixTabTargetV1(15_000, 15_000, 4_000)).toBe(16_000)
    expect(nativeDocxListSuffixTabTargetV1(15_000, 15_000, 0)).toBeUndefined()
    expect(nativeDocxListSuffixTabTargetV1(7_000, 15_000, 4_000, 21_000)).toBe(21_000)
    expect(nativeDocxListSuffixTabTargetV1(21_000, 15_000, 4_000, 21_000)).toBeUndefined()
  })

  it('pre-refuses every physical/logical alignment outside its exact LTR/RTL label boundary', () => {
    for (const direction of ['ltr', 'rtl'] as const) {
      for (const [alignment, boundary] of [['left', 10_000], ['right', 5_000], ['center', 10_000]] as const) {
        expect(positionNativeDocxListMarkerV1(marker(alignment), direction, boundary), `${direction} ${alignment} boundary`).toBeDefined()
        expect(positionNativeDocxListMarkerV1(marker(alignment), direction, boundary + 1), `${direction} ${alignment} overflow`).toBeUndefined()
      }
      const leading = direction === 'ltr' ? 'start' : 'end'
      const trailing = direction === 'ltr' ? 'end' : 'start'
      expect(positionNativeDocxListMarkerV1(marker(leading), direction, 10_000)).toBeDefined()
      expect(positionNativeDocxListMarkerV1(marker(leading), direction, 10_001)).toBeUndefined()
      expect(positionNativeDocxListMarkerV1(marker(trailing), direction, 5_000)).toBeDefined()
      expect(positionNativeDocxListMarkerV1(marker(trailing), direction, 5_001)).toBeUndefined()
    }
  })
})

/** The ideographic systems the Go resolver now renders have to survive the
 * wire: a format the validator does not know rejects the whole layout, so the
 * enum is the only thing standing between a resolved CJK marker and a refusal
 * that names nothing. */
describe('native DOCX ideographic marker formats on the resolved-layout wire', () => {
  function layout(format: string): unknown {
    return {
      protocol: 'injoffice.docx.resolved-layout',
      version: 1,
      document_id: 'document:1',
      revision: 'rev:1',
      source_parts: { main_part: 'word/document.xml' },
      paragraphs: [{
        paragraph_id: 'paragraph:1',
        applied_styles: [],
        properties: {},
        paragraph_mark_properties: { font_family: 'Test', font_size_half_points: 20 },
        numbering: {
          marker_id: 'marker:1', definition_sha256: `sha256:${'a'.repeat(64)}`, num_id: '1', abstract_num_id: '1',
          level: 0, start: 1, format, text: '%1.', suffix: 'tab', alignment: 'left', never_restart: true,
          counter_value: 10, counter_values: [{ level: 0, value: 10, format }], resolved_text: '一零.',
          label_start_twips: 0, label_end_twips: 480, text_start_twips: 480, marker_properties: {},
        },
      }],
      runs: [{ run_id: 'run:1', paragraph_id: 'paragraph:1', applied_paragraph_styles: [], applied_character_styles: [], properties: { font_family: 'Test', font_size_half_points: 20 } }],
      tables: [], fonts: [{ name: 'Test' }], diagnostics: [],
    }
  }
  it('accepts every modeled system and still rejects one that is not modeled', () => {
    for (const format of ['ideographTraditional', 'ideographZodiac', 'ideographLegalTraditional', 'taiwaneseCountingThousand', 'koreanDigital2']) {
      expect(decodeNativeDocxResolvedLayout(layout(format)), format).toMatchObject({ ok: true })
    }
    expect(decodeNativeDocxResolvedLayout(layout('ideographDigital'))).toMatchObject({ ok: false })
  })
})
