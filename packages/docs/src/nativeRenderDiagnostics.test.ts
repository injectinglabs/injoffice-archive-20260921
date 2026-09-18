import { describe, expect, it } from 'vitest'
import { isRenderNeutralLayoutDiagnostic } from './nativeRenderDiagnostics.js'
import { DOCX_RESOLVED_LAYOUT_PROTOCOL, DOCX_RESOLVED_LAYOUT_VERSION, type NativeDocxResolutionDiagnosticV1, type NativeDocxResolvedLayoutInputV1 } from './nativeResolvedLayout.js'

const resolved: NativeDocxResolvedLayoutInputV1 = {
  protocol: DOCX_RESOLVED_LAYOUT_PROTOCOL, version: DOCX_RESOLVED_LAYOUT_VERSION,
  document_id: 'document:styles', revision: 'revision:1',
  source_parts: { main_part: 'word/document.xml', styles_part: 'word/styles.xml' },
  paragraphs: [], runs: [], tables: [], fonts: [], diagnostics: [],
}

/**
 * ECMA-376 17.7.2: a `w:pStyle` or `w:rStyle` naming a style the package does
 * not define is ignored, so the consumer cascades from the document defaults
 * and no style layer was dropped. `floating-table-section-columns.docx` and
 * `StyleRef-DE.docx` in hard-v2 both reference a paragraph style `plain` their
 * styles part never defines, and Microsoft Word 16.112.4 paints both.
 */
describe('native DOCX render-neutral layout diagnostics', () => {
  const undefinedReference: NativeDocxResolutionDiagnosticV1 = {
    code: 'UNDEFINED_STYLE_REFERENCE', severity: 'unsupported', scope_id: 'paragraph:1',
    part_name: 'word/styles.xml', preservation: 'preserve-verbatim',
    message: 'The referenced paragraph style is not defined in this package; the reference is ignored and no style layer was dropped',
  }

  it('reads an undefined paragraph or character style reference as render-neutral', () => {
    expect(isRenderNeutralLayoutDiagnostic(undefinedReference, resolved)).toBe(true)
  })

  it('exempts only the resolver own shape of that disclosure', () => {
    // The severity and preservation branches are defensive: the wire type
    // admits only one value of each, so a wider value can only arrive from a
    // decoder that was bypassed, and the gate must still reject it.
    const variants: NativeDocxResolutionDiagnosticV1[] = [
      { ...undefinedReference, severity: 'deferred' } as unknown as NativeDocxResolutionDiagnosticV1,
      { ...undefinedReference, preservation: 'omit' } as unknown as NativeDocxResolutionDiagnosticV1,
      { ...undefinedReference, part_name: 'word/numbering.xml' },
      { ...undefinedReference, part_name: undefined },
      { ...undefinedReference, path: '/w:styles[1]/w:style[1]' },
      { ...undefinedReference, code: 'UNKNOWN_STYLE_CODE' },
    ]
    for (const variant of variants) expect(isRenderNeutralLayoutDiagnostic(variant, resolved), JSON.stringify(variant)).toBe(false)
  })

  it('keeps a missing basedOn ancestor of a style that does exist blocking', () => {
    // A dropped layer is a different fact from an ignored reference, and only
    // an unresolvable table style head pairs its way out of this gate.
    expect(isRenderNeutralLayoutDiagnostic({
      ...undefinedReference, code: 'MISSING_STYLE_REFERENCE',
      message: 'The missing basedOn ancestor was ignored; available descendant layers were retained',
    }, resolved)).toBe(false)
  })
})
