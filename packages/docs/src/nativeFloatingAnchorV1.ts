import type { NativeDocxQualifiedInlineImageV1 } from './nativeImagePagePaintV1.js'
import type { NativeDocxPaginatedLayoutV1, NativeDocxPaginatedPageV1 } from './nativePaginationV1.js'

/** Where one anchoring paragraph landed, which is all a column/paragraph-relative
 * anchor needs to become a page coordinate. */
export interface NativeDocxFloatingAnchorOriginV1 {
  page_id: string
  /** Top of the paragraph's first placed line on that page. */
  paragraph_top_millipoints: number
}

export interface NativeDocxResolvedFloatingAnchorV1 {
  x_millipoints: number
  y_millipoints: number
  /** The wrap region: the drawing box widened by distL/distR. */
  exclusion_left_millipoints: number
  exclusion_right_millipoints: number
}

/**
 * MS-DOCX anchors a floating drawing to a box that is not always the page. A
 * `page` origin is already absolute. A `margin` origin, and a `column` origin in
 * the single-column bodies this slice admits, are the body box's left edge. A
 * `paragraph` origin is the top of the anchoring paragraph, which only exists
 * once provisional pagination has placed it — so every caller resolves after
 * pagination, never inside shaping.
 */
export function resolveNativeDocxFloatingAnchorV1(
  floating: NonNullable<NativeDocxQualifiedInlineImageV1['floating']>,
  widthMilliPoints: number,
  page: NativeDocxPaginatedPageV1,
  origin: NativeDocxFloatingAnchorOriginV1,
): NativeDocxResolvedFloatingAnchorV1 {
  if (page.columns.length !== 1) throw new Error('Floating anchor origins require a single-column section')
  if (origin.page_id !== page.id) throw new Error('Floating anchor resolved against a page that does not own its paragraph')
  const horizontal = floating.horizontal_origin === 'page' ? 0 : page.body_box.x_millipoints
  const vertical = floating.vertical_origin === 'page' ? 0 : origin.paragraph_top_millipoints
  const x = horizontal + floating.offset_x_millipoints
  const y = vertical + floating.offset_y_millipoints
  if (!Number.isSafeInteger(x) || !Number.isSafeInteger(y)) throw new Error('Resolved floating anchor is not an exact milli-point coordinate')
  return {
    x_millipoints: x,
    y_millipoints: y,
    exclusion_left_millipoints: x - floating.wrap_distance_left_millipoints,
    exclusion_right_millipoints: x + widthMilliPoints + floating.wrap_distance_right_millipoints,
  }
}

/**
 * Index every body paragraph by where pagination placed it. A paragraph whose
 * lines straddle a page break has no single origin, so it is left out and its
 * anchors refuse rather than guess a page.
 */
export function nativeDocxFloatingAnchorOriginsV1(layout: NativeDocxPaginatedLayoutV1): Map<string, NativeDocxFloatingAnchorOriginV1> {
  const origins = new Map<string, NativeDocxFloatingAnchorOriginV1>()
  const straddled = new Set<string>()
  for (const page of layout.pages) for (const placed of page.lines) {
    const existing = origins.get(placed.paragraph_id)
    if (existing === undefined) {
      origins.set(placed.paragraph_id, { page_id: page.id, paragraph_top_millipoints: placed.y_millipoints })
      continue
    }
    if (existing.page_id !== page.id) { straddled.add(placed.paragraph_id); continue }
    if (placed.y_millipoints < existing.paragraph_top_millipoints) existing.paragraph_top_millipoints = placed.y_millipoints
  }
  for (const paragraph of straddled) origins.delete(paragraph)
  return origins
}
