import { describe, expect, it } from 'vitest'
import { PPTX_GROUP_LOCKS_PRESERVED_CODE } from './groupLocks'
import type { NativeDiagnostic, NativeElement, NativePptxDeck } from './types'
import { validateNativePptx } from './validate'

const preserved: NativeDiagnostic = {
  severity: 'warning', code: PPTX_GROUP_LOCKS_PRESERVED_CODE,
  message: 'group a:grpSpLocks editing locks are source-preserved but not modeled in native PPTX v1',
}

function leaf(): NativeElement {
  return {
    kind: 'shape', id: 'leaf', provenance: 'authored', preset: 'rect',
    transform: { x: 0, y: 0, cx: 10_000, cy: 10_000 },
    paragraphs: [], passthrough: [], compatibility: { status: 'preserveOnly', diagnostics: [] },
  }
}

function group(diagnostics: NativeDiagnostic[], status: 'editable' | 'preserveOnly'): NativeElement {
  return {
    kind: 'group', id: 'locked-group', provenance: 'authored',
    transform: { x: 0, y: 0, cx: 100_000, cy: 100_000 },
    childTransform: { x: 0, y: 0, cx: 100_000, cy: 100_000 },
    children: [leaf()], passthrough: [], compatibility: { status, diagnostics },
  }
}

function deck(elements: NativeElement[]): NativePptxDeck {
  return {
    contractVersion: 'pptx-native/v1', documentId: 'group-locks-deck', origin: 'authored',
    size: { cx: 2_000_000, cy: 1_500_000 }, assets: [],
    slides: [{ id: 'slide-1', provenance: 'authored', elements, passthrough: [], compatibility: { status: 'preserveOnly', diagnostics: [] } }],
    compatibility: { status: 'preserveOnly', diagnostics: [] },
  }
}

describe('preserved group lock contract', () => {
  it('pins the diagnostic code shared with the Go extractor', () => {
    expect(PPTX_GROUP_LOCKS_PRESERVED_CODE).toBe('pptx.group-locks-preserved')
  })

  it('rejects the preserved lock code on an editable group', () => {
    const result = validateNativePptx(deck([group([preserved], 'editable')]))
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.issues.map((issue) => issue.code)).toContain('native.groupLocksPreserved')
  })

  it('rejects the preserved lock code on a non-group element', () => {
    const text: NativeElement = {
      kind: 'text', id: 'text-lock', provenance: 'authored', transform: { x: 0, y: 0, cx: 100_000, cy: 100_000 },
      paragraphs: [{ align: 'left', level: 0, bullet: false, runs: [{ text: 'A' }] }],
      passthrough: [], compatibility: { status: 'preserveOnly', diagnostics: [preserved] },
    }
    const result = validateNativePptx(deck([text]))
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.issues.map((issue) => issue.code)).toContain('native.groupLocksPreserved')
  })
})
