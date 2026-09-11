import type { NativeDocxDocumentV1 } from './nativeContract.js'
import type { NativeDocxPaginatedLayoutV1 } from './nativePaginationV1.js'

/** Decimal display numbers follow physical pages, with explicit section restarts. */
export function nativeDocxPageNumberV1(document: NativeDocxDocumentV1, layout: NativeDocxPaginatedLayoutV1, ordinal: number): number {
  if (layout.status !== 'paginated' || !Number.isInteger(ordinal) || ordinal < 0 || ordinal >= layout.pages.length) throw new TypeError('Page numbering requires a final page identity')
  const sections = new Map(document.sections.map(section => [section.id, section]))
  const seen = new Set<string>()
  let number = 0
  for (const page of layout.pages.slice(0, ordinal + 1)) {
    number++
    for (const id of page.section_ids) {
      const section = sections.get(id)
      if (!section) throw new TypeError('Page numbering requires exact section identities')
      if (!seen.has(id)) {
        if (section.page_number_start !== undefined) {
          if (page.section_ids.length !== 1 || page.kind !== 'content') throw new TypeError('Decimal restart on a shared continuous-section or parity page remains unsupported')
          number = section.page_number_start
        }
        seen.add(id)
      }
    }
    if (number < 0 || number > 999999) throw new RangeError('Displayed decimal page number exceeds bounded profile')
  }
  return number
}
