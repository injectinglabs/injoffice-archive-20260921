import { describe, expect, it } from 'vitest'
import type { NativeDocxEditPolicyV1, NativeDocxTableV1 } from './nativeContract.js'
import type { NativeDocxResolvedLayoutInputV1 } from './nativeResolvedLayout.js'
import type { NativeDocxShapedLinesV1 } from './nativeShapingLines.js'
import { resolveNativeDocxTableAutofitV1 } from './nativeTableAutofitV1.js'

const HASH = `sha256:${'a'.repeat(64)}`
const anchor = (path: string) => ({ part_name: 'word/document.xml', path, start_byte: 1, end_byte: 2, xml_sha256: HASH })
const policy: NativeDocxEditPolicyV1 = { mode: 'read-only', allowed_operations: [], refusal: { code: 'NATIVE_READ_ONLY', message: 'Source remains authoritative.', preservation: 'refuse-mutation' } }

function shapedParagraph(id: string, advance = 0): NativeDocxShapedLinesV1['paragraphs'][number] {
  return {
    paragraph_id: id, story_id: 'story:body', story_kind: 'body', direction: 'ltr', alignment: 'start',
    spacing_before_millipoints: 0, spacing_after_millipoints: 0, indent_start_millipoints: 0, indent_end_millipoints: 0, first_line_delta_millipoints: 0,
    block_advance_millipoints: 6_000,
    lines: [{
      id: `line:${id}:0`, ordinal: 0, available_width_millipoints: 1_000_000_000, inline_offset_millipoints: 0, advance_inline_millipoints: advance,
      ascent_millipoints: 5_000, descent_millipoints: -1_000, line_gap_millipoints: 0, line_height_millipoints: 6_000, justified: false, logical_to_visual: advance ? [0] : [],
      fragments: advance ? [{
        id: `fragment:${id}:0:0`, source_kind: 'run', source_id: `${id}:run`, start_utf16: 0, end_utf16: 1, text: '1', direction: 'ltr', bidi_level: 0, logical_order: 0,
        script: 'Latn', language: 'fr-FR', whitespace: false, advance_inline_millipoints: advance, justification_expansion_millipoints: 0,
        ascent_millipoints: 5_000, descent_millipoints: -1_000, line_gap_millipoints: 0, glyphs: [],
      }] : [],
    }],
  }
}

function tableGrid(): { table: NativeDocxTableV1; resolved: NativeDocxResolvedLayoutInputV1; shaped: NativeDocxShapedLinesV1 } {
  const cell = (id: string, width: number, text?: string): NativeDocxTableV1['rows'][number]['cells'][number] => ({
    id, anchor: anchor(`/w:tc[${id}]`), width_twips: width, grid_span: 1, vertical_merge: 'none',
    paragraphs: [{ id: `${id}:p`, anchor: anchor(`/w:p[${id}]`), edit_policy: policy, properties: {}, runs: text ? [{ id: `${id}:p:run`, anchor: anchor(`/w:r[${id}]`), kind: 'text', text }] : [] }],
  })
  const table: NativeDocxTableV1 = {
    id: 'table:1', anchor: anchor('/w:tbl[1]'), edit_policy: policy, table_style_id: 'TableGrid', layout: 'autofit',
    indent_twips: 0, grid_widths_twips: [1_510, 1_511],
    cell_margins: { top_twips: 0, right_twips: 108, bottom_twips: 0, left_twips: 108 },
    rows: [
      { id: 'row:1', anchor: anchor('/w:tr[1]'), repeat_header: false, cant_split: false, cells: [cell('cell:1', 1_510, '1'), cell('cell:2', 1_511)] },
    ],
  }
  const resolved: NativeDocxResolvedLayoutInputV1 = {
    protocol: 'injoffice.docx.resolved-layout', version: 1, document_id: 'document:table', revision: 'revision:1',
    source_parts: { main_part: 'word/document.xml' },
    paragraphs: [{ paragraph_id: 'cell:1:p', applied_styles: [], properties: {}, paragraph_mark_properties: {} }, { paragraph_id: 'cell:2:p', applied_styles: [], properties: {}, paragraph_mark_properties: {} }],
    runs: [], tables: [{ table_id: table.id }], fonts: [], diagnostics: [],
  }
  const shaped: NativeDocxShapedLinesV1 = {
    protocol: 'injoffice.docx.shaped-lines', version: 1, document_id: 'document:table', revision: 'revision:1',
    available_width_millipoints: 453_600, tab_interval_millipoints: 36_000,
    font_manifest: { manifest_id: 'manifest:test', revision: 'revision:1' },
    providers: { resolver_id: 'r', resolver_revision: '1', shaper_id: 's', shaper_revision: '1', bidi_id: 'b', bidi_revision: '1', bidi_unicode_version: '13.0.0', unicode13_revision: '1' },
    diagnostics: [], paragraphs: [shapedParagraph('cell:1:p', 5_000), shapedParagraph('cell:2:p')],
  }
  return { table, resolved, shaped }
}

describe('source-preferred TableGrid autofit', () => {
  it('keeps authored grid columns when empty cells and short digits fit', () => {
    const { table, resolved, shaped } = tableGrid()
    const original = JSON.stringify(table)
    const result = resolveNativeDocxTableAutofitV1(table, 9_072, 'section:1', resolved, shaped)
    expect(result).toMatchObject({
      policy: { name: 'source-preferred-nonconflicting-v1', source_grid_widths_twips: [1_510, 1_511], preferred_width_twips: null },
      table: { layout: 'fixed', width_twips: 3_021, grid_widths_twips: [1_510, 1_511] },
    })
    expect(JSON.stringify(table)).toBe(original)
  })

  it('refuses merges, percent widths, mixed-space fragments, and unsatisfied minima', () => {
    for (const mode of ['merge', 'percent', 'space', 'wide'] as const) {
      const { table, resolved, shaped } = tableGrid()
      if (mode === 'merge') table.rows[0]!.cells[0]!.vertical_merge = 'restart'
      if (mode === 'percent') table.width_percent_fiftieths = 2_500
      if (mode === 'space') shaped.paragraphs[0]!.lines[0]!.fragments[0]!.text = '1 2'
      if (mode === 'wide') shaped.paragraphs[0]!.lines[0]!.fragments[0]!.advance_inline_millipoints = 1_000_000_000
      expect(resolveNativeDocxTableAutofitV1(table, 9_072, 'section:1', resolved, shaped)).toBeUndefined()
    }
  })
})
