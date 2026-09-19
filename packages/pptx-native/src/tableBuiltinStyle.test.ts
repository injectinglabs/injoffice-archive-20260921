import { describe, expect, it } from 'vitest'
import { PPTX_TABLE_BUILTIN_STYLE_POLICY, PPTX_TABLE_BUILTIN_STYLE_PREVIEW_CODE, PPTX_TABLE_NONVISUAL_PRESERVED_CODE } from './tableBuiltinStyle'
import type { NativeDiagnostic, NativeElement, NativePptxDeck } from './types'
import { validateNativePptx } from './validate'

function styledCell(fill: string) {
  return { text: '', fill, border: { color: 'FFFFFF', widthEmu: 12_700 } }
}

function styledTable(diagnostics: NativeDiagnostic[], status: 'editable' | 'preserveOnly' = 'editable'): NativeElement {
  return {
    kind: 'table', id: 'styled-table', provenance: 'authored',
    transform: { x: 100_000, y: 200_000, cx: 400_000, cy: 200_000 },
    table: { columnWidths: [200_000, 200_000], rowHeights: [200_000], rows: [[styledCell('000000'), styledCell('CBCBCB')]] },
    passthrough: [], compatibility: { status, diagnostics },
  }
}

function deck(elements: NativeElement[]): NativePptxDeck {
  return {
    contractVersion: 'pptx-native/v1', documentId: 'table-style-preview-deck', origin: 'authored',
    size: { cx: 2_000_000, cy: 1_500_000 }, assets: [],
    slides: [{ id: 'slide-1', provenance: 'authored', elements, passthrough: [], compatibility: { status: 'editable', diagnostics: [] } }],
    compatibility: { status: 'editable', diagnostics: [] },
  }
}

const previewWarning: NativeDiagnostic = {
  severity: 'warning', code: PPTX_TABLE_BUILTIN_STYLE_PREVIEW_CODE,
  message: `built-in table style resolved with policy ${PPTX_TABLE_BUILTIN_STYLE_POLICY}`,
}

describe('built-in table style preview contract', () => {
  it('pins the diagnostic code and policy string shared with the Go extractor', () => {
    expect(PPTX_TABLE_BUILTIN_STYLE_PREVIEW_CODE).toBe('pptx.table-builtin-style-preview')
    expect(PPTX_TABLE_BUILTIN_STYLE_POLICY).toBe('builtin-table-style-catalog-linear-srgb-tint-v1')
  })

  it('accepts text-free legacy cells with fills and uniform borders', () => {
    const result = validateNativePptx(deck([styledTable([])]))
    expect(result.ok).toBe(true)
  })

  it('rejects the preview code on an editable or unanchored table', () => {
    const result = validateNativePptx(deck([styledTable([previewWarning])]))
    expect(result.ok).toBe(false)
    if (!result.ok) {
      expect(result.issues.map((issue) => issue.code)).toContain('native.tableStylePreview')
      expect(result.issues.find((issue) => issue.code === 'native.tableStylePreview')?.path).toBe('$.slides[0].elements[0].compatibility')
    }
  })

  it('rejects the preserved lock/modId code on an editable or non-table element', () => {
    const preserved: NativeDiagnostic = {
      severity: 'warning', code: PPTX_TABLE_NONVISUAL_PRESERVED_CODE,
      message: 'table a:graphicFrameLocks and PowerPoint modId metadata are source-preserved but not modeled in native PPTX v1',
    }
    expect(PPTX_TABLE_NONVISUAL_PRESERVED_CODE).toBe('pptx.table-nonvisual-preserved')
    const editable = validateNativePptx(deck([styledTable([preserved])]))
    expect(editable.ok).toBe(false)
    if (!editable.ok) expect(editable.issues.map((issue) => issue.code)).toContain('native.tableNonVisualPreserved')
    const text: NativeElement = {
      kind: 'text', id: 'text-lock', provenance: 'authored', transform: { x: 0, y: 0, cx: 100_000, cy: 100_000 },
      paragraphs: [{ align: 'left', level: 0, bullet: false, runs: [{ text: 'A' }] }],
      passthrough: [], compatibility: { status: 'preserveOnly', diagnostics: [preserved] },
    }
    const wrongKind = validateNativePptx(deck([text]))
    expect(wrongKind.ok).toBe(false)
    if (!wrongKind.ok) expect(wrongKind.issues.map((issue) => issue.code)).toContain('native.tableNonVisualPreserved')
  })

  it('rejects the preview code on a non-table element even when read-only', () => {
    const text: NativeElement = {
      kind: 'text', id: 'text-1', provenance: 'authored', transform: { x: 0, y: 0, cx: 100_000, cy: 100_000 },
      paragraphs: [{ align: 'left', level: 0, bullet: false, runs: [{ text: 'A' }] }],
      passthrough: [], compatibility: { status: 'preserveOnly', diagnostics: [previewWarning] },
    }
    const result = validateNativePptx(deck([text]))
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.issues.map((issue) => issue.code)).toContain('native.tableStylePreview')
  })
})
