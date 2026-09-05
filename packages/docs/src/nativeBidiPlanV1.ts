/** Shared authoritative DOCX paragraph-to-UAX-#9 projection. */

import {
  NATIVE_BIDI_LIMITS,
  resolveNativeBidiParagraphV1,
  type NativeBidiExplicitRangeV1,
  type NativeBidiParagraphV1,
} from '@injoffice/font-metrics/bidi'
import type { NativeDocxParagraphV1 } from './nativeContract.js'
import type { NativeDocxResolvedRunV1 } from './nativeResolvedLayout.js'

export interface NativeDocxParagraphBidiPlanV1 {
  paragraph: NativeBidiParagraphV1
  runStarts: ReadonlyMap<string, number>
}

export type NativeDocxParagraphBidiPlanResultV1 =
  | { ok: true; value: NativeDocxParagraphBidiPlanV1 }
  | { ok: false; code: string; message: string; sourceID?: string }

function appendExplicitRanges(ranges: NativeBidiExplicitRangeV1[], text: string, start: number, direction: 'ltr' | 'rtl'): boolean {
  let chunkStart = 0
  for (let index = 0; index <= text.length; index++) {
    const separator = index === text.length || text.charCodeAt(index) === 0x0a || text.charCodeAt(index) === 0x0d
    if (!separator) continue
    if (index > chunkStart) {
      if (ranges.length >= NATIVE_BIDI_LIMITS.maxExplicitRanges) return false
      ranges.push({ startUtf16: start + chunkStart, endUtf16: start + index, direction })
    }
    chunkStart = index + 1
  }
  return true
}

export function resolveNativeDocxParagraphBidiPlanV1(
  paragraph: NativeDocxParagraphV1,
  resolvedRuns: ReadonlyMap<string, NativeDocxResolvedRunV1>,
  baseDirection: 'ltr' | 'rtl',
  referenceText: ReadonlyMap<string, string> = new Map(),
): NativeDocxParagraphBidiPlanResultV1 {
  let text = ''
  const runStarts = new Map<string, number>()
  const explicitRanges: NativeBidiExplicitRangeV1[] = []
  for (const run of paragraph.runs) {
    runStarts.set(run.id, text.length)
    const resolved = resolvedRuns.get(run.id)
    if (!resolved) return { ok: false, code: 'invalid-source', message: 'resolved run is missing', sourceID: run.id }
    if (resolved.properties.hidden) continue
    if (run.kind === 'text' && run.text !== undefined) {
      if (run.text.length > NATIVE_BIDI_LIMITS.maxUtf16 - text.length) return { ok: false, code: 'resource-limit', message: `bidi text exceeds ${NATIVE_BIDI_LIMITS.maxUtf16} UTF-16 code units`, sourceID: run.id }
      const start = text.length
      text += run.text
      if (resolved.properties.rtl !== undefined && !appendExplicitRanges(explicitRanges, run.text, start, resolved.properties.rtl ? 'rtl' : 'ltr')) return { ok: false, code: 'resource-limit', message: `explicit bidi ranges exceed ${NATIVE_BIDI_LIMITS.maxExplicitRanges}`, sourceID: run.id }
    } else if (run.kind === 'reference' && referenceText.has(run.id)) {
      const marker = referenceText.get(run.id)!
      if (marker.length > NATIVE_BIDI_LIMITS.maxUtf16 - text.length) return { ok: false, code: 'resource-limit', message: `bidi text exceeds ${NATIVE_BIDI_LIMITS.maxUtf16} UTF-16 code units`, sourceID: run.id }
      const start = text.length
      text += marker
      if (resolved.properties.rtl !== undefined && !appendExplicitRanges(explicitRanges, marker, start, resolved.properties.rtl ? 'rtl' : 'ltr')) return { ok: false, code: 'resource-limit', message: `explicit bidi ranges exceed ${NATIVE_BIDI_LIMITS.maxExplicitRanges}`, sourceID: run.id }
    } else if (run.kind === 'drawing') {
      if (text.length >= NATIVE_BIDI_LIMITS.maxUtf16) return { ok: false, code: 'resource-limit', message: `bidi text exceeds ${NATIVE_BIDI_LIMITS.maxUtf16} UTF-16 code units`, sourceID: run.id }
      const start = text.length
      text += '\uFFFC'
      if (resolved.properties.rtl !== undefined && !appendExplicitRanges(explicitRanges, '\uFFFC', start, resolved.properties.rtl ? 'rtl' : 'ltr')) return { ok: false, code: 'resource-limit', message: `explicit bidi ranges exceed ${NATIVE_BIDI_LIMITS.maxExplicitRanges}`, sourceID: run.id }
    } else if (run.kind === 'control') {
      const character = run.control === 'line-break' ? '\n' : run.control === 'tab' ? '\t' : run.control === 'soft-hyphen' ? '\u00ad' : ''
      if (character) {
        if (text.length >= NATIVE_BIDI_LIMITS.maxUtf16) return { ok: false, code: 'resource-limit', message: `bidi text exceeds ${NATIVE_BIDI_LIMITS.maxUtf16} UTF-16 code units`, sourceID: run.id }
        const start = text.length
        text += character
        if (resolved.properties.rtl !== undefined && !appendExplicitRanges(explicitRanges, character, start, resolved.properties.rtl ? 'rtl' : 'ltr')) return { ok: false, code: 'resource-limit', message: `explicit bidi ranges exceed ${NATIVE_BIDI_LIMITS.maxExplicitRanges}`, sourceID: run.id }
      }
    }
  }
  const result = resolveNativeBidiParagraphV1({ text, baseDirection, ...(explicitRanges.length > 0 ? { explicitRanges } : {}) })
  if (!result.ok) return { ok: false, code: result.code, message: result.message }
  return { ok: true, value: { paragraph: result.value, runStarts } }
}
