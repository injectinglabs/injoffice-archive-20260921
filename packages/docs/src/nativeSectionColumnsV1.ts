/** Exact, renderer-neutral WordprocessingML section/column geometry. */

import type { NativeDocxHeaderFooterReferenceV1, NativeDocxSectionV1 } from './nativeContract.js'

export const DOCX_SECTION_COLUMN_LIMITS = Object.freeze({ maxColumns: 45 })
const MILLIPOINTS_PER_TWIP = 50
const MAX_COORDINATE = 1_000_000_000_000

export interface NativeDocxQualifiedColumnV1 {
  id: string
  ordinal: number
  x_millipoints: number
  y_millipoints: number
  width_millipoints: number
  height_millipoints: number
}

export interface NativeDocxQualifiedSectionGeometryV1 {
  page_width_millipoints: number
  page_height_millipoints: number
  body_x_millipoints: number
  body_y_millipoints: number
  body_width_millipoints: number
  body_height_millipoints: number
  columns: NativeDocxQualifiedColumnV1[]
}

export type QualifyNativeDocxSectionColumnsV1Result =
  | { ok: true; value: NativeDocxQualifiedSectionGeometryV1 }
  | { ok: false; code: 'section-geometry-invalid' | 'column-geometry-invalid' | 'unequal-column-widths'; message: string }

function twips(value: number): number | undefined {
  const converted = value * MILLIPOINTS_PER_TWIP
  return Number.isSafeInteger(converted) && converted >= 0 && converted <= MAX_COORDINATE ? converted : undefined
}

function checkedSum(...values: number[]): number | undefined {
  let sum = 0
  for (const value of values) {
    sum += value
    if (!Number.isSafeInteger(sum) || Math.abs(sum) > MAX_COORDINATE) return undefined
  }
  return sum
}

export function qualifyNativeDocxSectionColumnsV1(section: NativeDocxSectionV1, options?: { allowUnequalWidths: true }): QualifyNativeDocxSectionColumnsV1Result {
  const pageWidth = twips(section.page.width_twips)
  const pageHeight = twips(section.page.height_twips)
  const left = twips(section.page.margins.left_twips)
  const right = twips(section.page.margins.right_twips)
  const top = twips(section.page.margins.top_twips)
  const bottom = twips(section.page.margins.bottom_twips)
  const gutter = twips(section.page.margins.gutter_twips)
  if ([pageWidth, pageHeight, left, right, top, bottom, gutter].some((value) => value === undefined)) {
    return { ok: false, code: 'section-geometry-invalid', message: 'Section dimensions exceed bounded integer milli-point geometry' }
  }
  if ((section.page.orientation === 'portrait' && pageWidth! > pageHeight!) || (section.page.orientation === 'landscape' && pageWidth! < pageHeight!)) {
    return { ok: false, code: 'section-geometry-invalid', message: 'Section dimensions contradict the declared orientation' }
  }
  // ECMA-376 17.6.19: rtlGutter binds on the right, so the gutter is taken out
  // of the right margin and the body box starts at the plain left margin.
  const bodyX = section.page.rtl_gutter === true ? left! : checkedSum(left!, gutter!)
  const bodyWidth = checkedSum(pageWidth!, -left!, -right!, -gutter!)
  const bodyHeight = checkedSum(pageHeight!, -top!, -bottom!)
  if (bodyX === undefined || bodyWidth === undefined || bodyHeight === undefined || bodyWidth <= 0 || bodyHeight <= 0) {
    return { ok: false, code: 'section-geometry-invalid', message: 'Section margins leave a non-positive or unbounded body box' }
  }
  if (section.page.columns < 1 || section.page.columns > DOCX_SECTION_COLUMN_LIMITS.maxColumns || section.page.column_definitions.length !== section.page.columns) {
    return { ok: false, code: 'column-geometry-invalid', message: 'Column count and stable column identities must agree within the Word 1..45 bound' }
  }
  const widths: number[] = []
  const gaps: number[] = []
  if (section.page.column_layout === 'equal-width') {
    const bodyWidthTwips = section.page.width_twips - section.page.margins.left_twips - section.page.margins.right_twips - section.page.margins.gutter_twips
    const totalGapsTwips = section.page.column_spacing_twips * (section.page.columns - 1)
    const availableTwips = bodyWidthTwips - totalGapsTwips
    const gap = twips(section.page.column_spacing_twips)
    const totalGaps = gap === undefined ? undefined : checkedSum(...Array.from({ length: section.page.columns - 1 }, () => gap))
    const available = totalGaps === undefined ? undefined : checkedSum(bodyWidth, -totalGaps)
    if (!Number.isSafeInteger(availableTwips) || availableTwips <= 0 || availableTwips % section.page.columns !== 0 || gap === undefined || available === undefined || available <= 0) {
      return { ok: false, code: 'column-geometry-invalid', message: 'Equal-width column geometry is not exactly divisible in integer twips' }
    }
    for (let index = 0; index < section.page.columns; index += 1) {
      widths.push(available / section.page.columns)
      gaps.push(index + 1 === section.page.columns ? 0 : gap)
    }
  } else {
    for (const [index, definition] of section.page.column_definitions.entries()) {
      const width = definition.width_twips === undefined ? undefined : twips(definition.width_twips)
      const gap = definition.space_after_twips === undefined ? undefined : twips(definition.space_after_twips)
      if (width === undefined || width <= 0 || gap === undefined || (index + 1 === section.page.columns && gap !== 0)) {
        return { ok: false, code: 'column-geometry-invalid', message: 'Explicit columns require positive widths, bounded gaps, and zero trailing space' }
      }
      widths.push(width)
      gaps.push(gap)
    }
    let total = 0
    for (let index = 0; index < widths.length; index += 1) {
      const next = checkedSum(total, widths[index]!, gaps[index]!)
      if (next === undefined) return { ok: false, code: 'column-geometry-invalid', message: 'Explicit column geometry exceeds bounded coordinates' }
      total = next
    }
    if (total !== bodyWidth) return { ok: false, code: 'column-geometry-invalid', message: 'Explicit column widths and gaps must exactly cover the section body width' }
  }
  if (!options?.allowUnequalWidths && widths.some((width) => width !== widths[0])) {
    return { ok: false, code: 'unequal-column-widths', message: 'Unequal column widths require per-column reshaping and are outside the exact v1 slice' }
  }
  const columns: NativeDocxQualifiedColumnV1[] = []
  let offset = 0
  for (let index = 0; index < widths.length; index += 1) {
    const definition = section.page.column_definitions[index]!
    const x = checkedSum(bodyX, offset)
    if (x === undefined) return { ok: false, code: 'column-geometry-invalid', message: 'Column origin exceeds bounded coordinates' }
    columns.push({ id: definition.id, ordinal: index, x_millipoints: x, y_millipoints: top!, width_millipoints: widths[index]!, height_millipoints: bodyHeight })
    const next = checkedSum(offset, widths[index]!, gaps[index]!)
    if (next === undefined) return { ok: false, code: 'column-geometry-invalid', message: 'Column extent exceeds bounded coordinates' }
    offset = next
  }
  return {
    ok: true,
    value: {
      page_width_millipoints: pageWidth!, page_height_millipoints: pageHeight!,
      body_x_millipoints: bodyX, body_y_millipoints: top!, body_width_millipoints: bodyWidth, body_height_millipoints: bodyHeight,
      columns,
    },
  }
}

function sameRefs(left: readonly NativeDocxHeaderFooterReferenceV1[], right: readonly NativeDocxHeaderFooterReferenceV1[]): boolean {
  return left.length === right.length && left.every((value, index) => {
    const other = right[index]
    return other !== undefined && value.kind === other.kind && value.story_id === other.story_id && value.relationship_id === other.relationship_id
  })
}

/**
 * Continuous and next-column transitions share a page only across identical
 * physical grids and page-level references.
 *
 * `allow_different_columns` drops only the column-division comparison. A
 * continuous break is Word's way of changing the column division part-way down
 * a page — it is how an index is set in columns and then returned to one — so
 * the division is exactly what such a break is allowed to change.
 *
 * `allow_different_body_vertical` drops the `body_y`/`body_height`
 * comparison — the top and bottom margins. The physical sheet's vertical frame
 * belongs to the section that opened the page: ECMA-376 17.6.11 gives `w:pgMar`
 * per section, but a `continuous` break does not start a page, so the top and
 * bottom margins of the continuing section have no page of their own to apply
 * to and the shared sheet keeps the frame it was opened with. Only the
 * horizontal geometry — the body box's sides and the column division — is
 * expressed mid-page. `office-hard-v2/pdf/floating-table-section-columns.pdf`
 * paints exactly that: a `pgMar 720` two-column section and a `pgMar`
 * top/bottom 1440, left/right 1080 single-column continuing section share
 * **one** 612 x 792 page (`/Type /Pages … /Count 1`), and the continuing
 * section's line is anchored at x = 54 pt — its own 1080-twip left margin —
 * below the two-column content rather than at its own 72 pt body top.
 *
 * `allow_different_body_sides` additionally drops the left/right comparison —
 * the body box's origin and width — because a continuous break is equally
 * Word's way of indenting the remainder of a page. Only the sides move: the
 * page, the top of the body box and its height are still compared, because
 * those are what fixes the shared physical sheet.
 * `office-hard-v2/pdf/endingSectionProps.pdf` paints exactly that — the
 * continuing section's single line runs 128.7 pt to 466.6 pt, its own
 * 2574-twip margins, on the same page as a section whose margins are 1134.
 *
 * Header/footer references and the title-page policy belong to the section
 * that opened the shared page. A continuing section that declares no
 * references of its own inherits them (ECMA-376 17.10.1), so it cannot
 * contradict that page and is not compared against it; one that declares its
 * own must state the same page. In the same PDF the shared page carries
 * section 1's first-page footer although the continuing section sets no
 * `titlePg` and names no stories.
 */
export function nativeDocxSectionsShareExactPageV1(previous: NativeDocxSectionV1, next: NativeDocxSectionV1, options?: { allow_different_columns?: boolean; allow_different_body_sides?: boolean; allow_different_body_vertical?: boolean }): boolean {
  const left = qualifyNativeDocxSectionColumnsV1(previous)
  const right = qualifyNativeDocxSectionColumnsV1(next)
  if (!left.ok || !right.ok) return false
  const a = left.value
  const b = right.value
  const sameColumns = a.columns.length === b.columns.length && a.columns.every((column, index) => {
    const other = b.columns[index]
    return other !== undefined && column.x_millipoints === other.x_millipoints && column.y_millipoints === other.y_millipoints && column.width_millipoints === other.width_millipoints && column.height_millipoints === other.height_millipoints
  })
  const sameSides = a.body_x_millipoints === b.body_x_millipoints && a.body_width_millipoints === b.body_width_millipoints
  const sameVertical = a.body_y_millipoints === b.body_y_millipoints && a.body_height_millipoints === b.body_height_millipoints
  const inheritsStories = next.header_refs.length === 0 && next.footer_refs.length === 0
  const sameStories = inheritsStories || (previous.title_page === next.title_page && sameRefs(previous.header_refs, next.header_refs) && sameRefs(previous.footer_refs, next.footer_refs))
  return a.page_width_millipoints === b.page_width_millipoints && a.page_height_millipoints === b.page_height_millipoints &&
    (sameVertical || options?.allow_different_body_vertical === true) &&
    (sameSides || options?.allow_different_body_sides === true) &&
    previous.page.margins.header_twips === next.page.margins.header_twips && previous.page.margins.footer_twips === next.page.margins.footer_twips &&
    (sameColumns || options?.allow_different_columns === true) && sameStories
}
