import type { NativePptxDeck } from './types'
import { assertNativePptx } from './validate'
import { PPTX_NATIVE_RESOURCE_LIMITS } from './schema.generated'

/**
 * Canonical JSON for hashes, caches, tests, and revision comparison.
 *
 * Object keys are sorted by UTF-16 code unit. Arrays retain semantic order;
 * assets are sorted by durable id before encoding, while slides, elements,
 * paragraphs, rows, and runs are never reordered. The v1 contract admits
 * finite safe integers only, avoiding language-specific float formatting.
 */
export function canonicalizeNativePptx(deck: NativePptxDeck): NativePptxDeck {
  return {
    ...deck,
    assets: [...deck.assets].sort((left, right) => compareUtf16(left.id, right.id)),
  }
}

export function stringifyNativePptx(deck: NativePptxDeck): string {
  assertNativePptx(deck)
  const encoded = stringifyCanonicalValue(canonicalizeNativePptx(deck))
  if (utf8ByteLength(encoded) > PPTX_NATIVE_RESOURCE_LIMITS.maxJsonBytes) {
    throw new TypeError(`native PPTX JSON exceeds ${PPTX_NATIVE_RESOURCE_LIMITS.maxJsonBytes} UTF-8 bytes`)
  }
  return encoded
}

function stringifyCanonicalValue(value: unknown): string {
  if (value === null) return 'null'
  if (typeof value === 'string') return quote(value)
  if (typeof value === 'boolean') return value ? 'true' : 'false'
  if (typeof value === 'number') {
    if (!Number.isSafeInteger(value) || Object.is(value, -0)) throw new TypeError('native PPTX JSON numbers must be safe non-negative-zero integers')
    return String(value)
  }
  if (Array.isArray(value)) return `[${value.map(stringifyCanonicalValue).join(',')}]`
  if (typeof value === 'object') {
    const record = value as Record<string, unknown>
    const keys = Object.keys(record).filter((key) => record[key] !== undefined).sort(compareUtf16)
    return `{${keys.map((key) => `${quote(key)}:${stringifyCanonicalValue(record[key])}`).join(',')}}`
  }
  throw new TypeError(`native PPTX JSON cannot encode ${typeof value}`)
}

function quote(value: string): string {
  return JSON.stringify(value).replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029')
}

function compareUtf16(left: string, right: string): number {
  return left < right ? -1 : left > right ? 1 : 0
}

function utf8ByteLength(value: string): number {
  let bytes = 0
  for (let index = 0; index < value.length; index++) {
    const codeUnit = value.charCodeAt(index)
    if (codeUnit < 0x80) bytes++
    else if (codeUnit < 0x800) bytes += 2
    else if (codeUnit >= 0xd800 && codeUnit <= 0xdbff && index + 1 < value.length && value.charCodeAt(index + 1) >= 0xdc00 && value.charCodeAt(index + 1) <= 0xdfff) {
      bytes += 4
      index++
    } else bytes += 3
  }
  return bytes
}
