/** Hash-bound Unicode 13 classification shared by itemization and shaping. */

import { sha256 } from '@noble/hashes/sha2.js'
import { bytesToHex } from '@noble/hashes/utils.js'
import {
  UNICODE_13_GENERATOR_SHA256,
  UNICODE_13_PROJECTION_SHA256,
  UNICODE_13_TABLES_ENCODING,
  UNICODE_13_TABLES_ENCODING_SHA256,
  UNICODE_13_UCD_SOURCE_SHA256,
  UNICODE_13_VERSION,
} from './unicode13Tables.generated.js'

export {
  UNICODE_13_GENERATOR_SHA256,
  UNICODE_13_PROJECTION_SHA256,
  UNICODE_13_TABLES_ENCODING_SHA256,
  UNICODE_13_UCD_SOURCE_SHA256,
  UNICODE_13_VERSION,
}

export type Unicode13Script = 'Latn' | 'Cyrl' | 'Grek' | 'Arab' | 'Hebr' | 'Deva' | 'Hani' | 'Kana' | 'Hang' | 'Zinh' | 'Zyyy'
export type Unicode13Punctuation = 'open' | 'close' | 'none'
type Range = readonly [number, number]

interface EncodingV1 {
  scripts: Array<[Unicode13Script, string]>
  whiteSpace: string
  defaultIgnorable: string
  assigned: string
  punctuation: { Ps: string; Pi: string; Pe: string; Pf: string; Po: string }
  control: string
}

function digestText(value: string): `sha256:${string}` {
  return `sha256:${bytesToHex(sha256(new TextEncoder().encode(value)))}`
}

function decodeRanges(value: string): readonly Range[] {
  if (value.length === 0) return Object.freeze([])
  let previousEnd = -1
  return Object.freeze(value.split(',').map((encoded): Range => {
    const pieces = encoded.split('..')
    if (pieces.length < 1 || pieces.length > 2) throw new Error('malformed Unicode 13 range encoding')
    const start = Number.parseInt(pieces[0]!, 16)
    const end = Number.parseInt(pieces[1] ?? pieces[0]!, 16)
    if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start <= previousEnd || end < start || end > 0x10ffff) throw new Error('non-canonical Unicode 13 range encoding')
    previousEnd = end
    return Object.freeze([start, end])
  }))
}

function validScalar(value: number): boolean {
  return Number.isSafeInteger(value) && value >= 0 && value <= 0x10ffff && !(value >= 0xd800 && value <= 0xdfff)
}

function includes(ranges: readonly Range[], codePoint: number): boolean {
  if (!validScalar(codePoint)) return false
  let low = 0
  let high = ranges.length - 1
  while (low <= high) {
    const middle = (low + high) >>> 1
    const [start, end] = ranges[middle]!
    if (codePoint < start) high = middle - 1
    else if (codePoint > end) low = middle + 1
    else return true
  }
  return false
}

let decoded: EncodingV1 | undefined
let scripts: ReadonlyArray<readonly [Unicode13Script, readonly Range[]]> = []
let whiteSpace: readonly Range[] = []
let defaultIgnorable: readonly Range[] = []
let assigned: readonly Range[] = []
let punctuation: Readonly<Record<keyof EncodingV1['punctuation'], readonly Range[]>> = { Ps: [], Pi: [], Pe: [], Pf: [], Po: [] }
let control: readonly Range[] = []
let valid = digestText(UNICODE_13_TABLES_ENCODING) === UNICODE_13_TABLES_ENCODING_SHA256
try {
  decoded = JSON.parse(UNICODE_13_TABLES_ENCODING) as EncodingV1
  scripts = Object.freeze(decoded.scripts.map(([tag, ranges]) => Object.freeze([tag, decodeRanges(ranges)] as const)))
  whiteSpace = decodeRanges(decoded.whiteSpace)
  defaultIgnorable = decodeRanges(decoded.defaultIgnorable)
  assigned = decodeRanges(decoded.assigned)
  punctuation = Object.freeze({ Ps: decodeRanges(decoded.punctuation.Ps), Pi: decodeRanges(decoded.punctuation.Pi), Pe: decodeRanges(decoded.punctuation.Pe), Pf: decodeRanges(decoded.punctuation.Pf), Po: decodeRanges(decoded.punctuation.Po) })
  control = decodeRanges(decoded.control)
} catch {
  valid = false
}

export const UNICODE_13_TABLES_RUNTIME_MATCH = valid
export const UNICODE_13_CLASSIFIER_REVISION = digestText(JSON.stringify({
  version: UNICODE_13_VERSION,
  sources: UNICODE_13_UCD_SOURCE_SHA256,
  projection: UNICODE_13_PROJECTION_SHA256,
  generator: UNICODE_13_GENERATOR_SHA256,
  tables: UNICODE_13_TABLES_ENCODING_SHA256,
  policy: 'injoffice.unicode13-script-properties.v1',
}))

export function unicode13Script(codePoint: number): Unicode13Script | 'Other' | undefined {
  if (!valid || !validScalar(codePoint)) return undefined
  for (const [tag, ranges] of scripts) if (includes(ranges, codePoint)) return tag
  return 'Other'
}

export function isUnicode13WhiteSpace(codePoint: number): boolean {
  return valid && includes(whiteSpace, codePoint)
}

export function isUnicode13DefaultIgnorable(codePoint: number): boolean {
  return valid && includes(defaultIgnorable, codePoint)
}

export function isUnicode13Assigned(codePoint: number): boolean {
  return valid && includes(assigned, codePoint)
}

export function isUnicode13Control(codePoint: number): boolean {
  return valid && includes(control, codePoint)
}

export function unicode13Punctuation(codePoint: number): Unicode13Punctuation {
  if (!valid) return 'none'
  if (includes(punctuation.Ps, codePoint) || includes(punctuation.Pi, codePoint)) return 'open'
  return includes(punctuation.Pe, codePoint) || includes(punctuation.Pf, codePoint) || includes(punctuation.Po, codePoint) ? 'close' : 'none'
}

export function isUnicode13TextWhiteSpace(text: string): boolean {
  if (!valid || text.length === 0) return false
  for (const character of text) if (!isUnicode13WhiteSpace(character.codePointAt(0)!)) return false
  return true
}

export function isUnicode13BreakableWhiteSpace(text: string): boolean {
  if (!valid || text.length === 0) return false
  for (const character of text) {
    const codePoint = character.codePointAt(0)!
    if (!isUnicode13WhiteSpace(codePoint) || codePoint === 0x00a0 || codePoint === 0x0085 || codePoint === 0x2028 || codePoint === 0x2029 || codePoint === 0x202f) return false
  }
  return true
}
