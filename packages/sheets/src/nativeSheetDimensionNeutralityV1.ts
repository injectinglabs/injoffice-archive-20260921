/**
 * Read-only source evidence that a bounded-projection disclosure was raised by
 * worksheet markup ECMA-376 gives no role in a row height, a column width or a
 * merged rectangle.
 *
 * The bounded dimension projection reads sheetFormatPr, cols/col, row and
 * mergeCells. A worksheet routinely carries markup outside that vocabulary —
 * view state, outline levels, thick-edge flags, text-descent metadata,
 * markup-compatibility attributes — and the extractor discloses all of it. This
 * evidence separates "we did not model it" from "the dimensions may be wrong".
 *
 * It never reproduces what it clears, and it never clears a code it does not
 * list: zeroHeight and unqualified metrics keep refusing exactly as before.
 *
 * FOREIGN_WORKSHEET_MARKUP is cleared only for the one shape the Go tier
 * qualifies: a worksheet whose every foreign-namespace child is a
 * markup-compatibility AlternateContent holding nothing but a SpreadsheetML
 * <controls> block. Those are anchored form controls, and ECMA-376 gives
 * CT_ObjectAnchor no role in a row height, a column width or a merged
 * rectangle. Any other foreign child keeps the code refusing.
 */
export const NATIVE_SHEET_DIMENSION_NEUTRALITY_V1_POLICY = 'non-dimensional-worksheet-markup-v1' as const

/** The only disclosure codes this evidence may clear. */
export const NATIVE_SHEET_DIMENSION_NEUTRALITY_V1_CODES = Object.freeze([
  'COLS_ATTRIBUTES',
  'COLUMN_DIMENSION_EXTRAS',
  'FOREIGN_WORKSHEET_MARKUP',
  'ROW_DIMENSION_EXTRAS',
  'SHEET_FORMAT_EXTRAS',
  'SHEET_VIEW_GEOMETRY',
  'WORKSHEET_ATTRIBUTES',
] as const)

export type NativeSheetDimensionNeutralityCodeV1 = (typeof NATIVE_SHEET_DIMENSION_NEUTRALITY_V1_CODES)[number]

export interface NativeSheetDimensionNeutralityV1 {
  sheet_part: string
  policy: typeof NATIVE_SHEET_DIMENSION_NEUTRALITY_V1_POLICY
  codes: NativeSheetDimensionNeutralityCodeV1[]
  warnings: string[]
}

/** Bounded decoder for the optional non-dimensional-markup evidence. */
export function decodeNativeSheetDimensionNeutralityV1(value: unknown): NativeSheetDimensionNeutralityV1[] {
  const fail = (): never => {
    throw new TypeError('Invalid worksheet dimension neutrality evidence')
  }
  if (!Array.isArray(value) || value.length > 64) return fail()
  const seen = new Set<string>()
  return value.map((input) => {
    if (!input || typeof input !== 'object' || Array.isArray(input)) return fail()
    const entry = input as Record<string, unknown>
    const keys = Object.keys(entry)
    if (keys.length !== 4 || ['sheet_part', 'policy', 'codes', 'warnings'].some((key) => !Object.hasOwn(entry, key))) return fail()
    const part = entry.sheet_part
    if (
      typeof part !== 'string' ||
      !part ||
      part.length > 1024 ||
      part.startsWith('/') ||
      part.includes('\\') ||
      part.split('/').some((segment) => !segment || segment === '.' || segment === '..') ||
      // eslint-disable-next-line no-control-regex
      /[\x00-\x1f]/.test(part) ||
      seen.has(part)
    )
      return fail()
    seen.add(part)
    if (entry.policy !== NATIVE_SHEET_DIMENSION_NEUTRALITY_V1_POLICY) return fail()
    const codes = entry.codes
    if (!Array.isArray(codes) || codes.length > NATIVE_SHEET_DIMENSION_NEUTRALITY_V1_CODES.length) return fail()
    const owned = new Set<string>()
    for (const code of codes) {
      if (typeof code !== 'string' || !(NATIVE_SHEET_DIMENSION_NEUTRALITY_V1_CODES as readonly string[]).includes(code) || owned.has(code)) return fail()
      owned.add(code)
    }
    const warnings = entry.warnings
    if (!Array.isArray(warnings) || warnings.length < 1 || warnings.length > 4 || warnings.some((warning) => typeof warning !== 'string' || !warning || warning.length > 1024)) return fail()
    return Object.freeze({
      sheet_part: part,
      policy: NATIVE_SHEET_DIMENSION_NEUTRALITY_V1_POLICY,
      codes: Object.freeze([...codes] as NativeSheetDimensionNeutralityCodeV1[]) as NativeSheetDimensionNeutralityCodeV1[],
      warnings: Object.freeze([...(warnings as string[])]) as string[],
    }) as NativeSheetDimensionNeutralityV1
  })
}

/**
 * The codes cleared for one worksheet part. An absent, duplicated or
 * non-joining entry clears nothing.
 */
export function nativeSheetDimensionNeutralityCodesV1(
  entries: readonly NativeSheetDimensionNeutralityV1[] | undefined,
  sheetPart: string | undefined,
): ReadonlySet<string> {
  if (!entries || !sheetPart) return new Set<string>()
  const matches = entries.filter((entry) => entry.sheet_part === sheetPart)
  if (matches.length !== 1) return new Set<string>()
  return new Set<string>(matches[0]!.codes)
}
