import { decodeNativeDocxDocument, type NativeDocxDocumentV1, type NativeDocxRunV1 } from './nativeContract.js'
import type { NativeDocxPaginationRequestV1, NativeDocxPaginatedLayoutV1 } from './nativePaginationV1.js'
import { nativeDocxPageNumberV1 } from './nativePageNumbersV1.js'

export type NativeDocxBodyPageFieldValuesV1 = Record<string, string>

export function nativeDocxBodyPageFieldRunsV1(document: NativeDocxDocumentV1): NativeDocxRunV1[] {
  const fields: NativeDocxRunV1[] = []
  for (const story of [document.body, ...document.headers, ...document.footers, ...document.notes, ...document.comment_stories]) for (const block of story.blocks) {
    const paragraphs = block.paragraph ? [block.paragraph] : block.table?.rows.flatMap(row => row.cells.flatMap(cell => cell.paragraphs)) ?? []
    if (paragraphs.some(paragraph => paragraph.runs.some(run => run.layout_page_field))) throw new TypeError('Original sources cannot contain internal layout field markers')
  }
  for (const block of document.body.blocks) {
    if (block.table?.rows.some(row => row.cells.some(cell => cell.paragraphs.some(p => p.runs.some(r => r.page_field))))) throw new TypeError('Body page fields inside tables remain unsupported')
    for (const run of block.paragraph?.runs ?? []) if (run.page_field) {
      if (run.kind !== 'text' || run.text !== '' || block.paragraph?.edit_policy.mode !== 'read-only') throw new TypeError('Body field sources require empty cached text and read-only paragraphs')
      fields.push(run)
      if (fields.length > 128) throw new RangeError('Body page fields exceed 128-run limit')
    }
  }
  return fields
}

export function nativeDocxBodyPageFieldDocumentV1(source: NativeDocxDocumentV1, values: NativeDocxBodyPageFieldValuesV1): NativeDocxDocumentV1 {
  const fields = nativeDocxBodyPageFieldRunsV1(source)
  if (Object.keys(values).length !== fields.length || fields.some(run => !/^(0|[1-9][0-9]{0,5})$/.test(values[run.id] ?? ''))) throw new TypeError('Body field values must exactly cover source fields with bounded decimal text')
  const derived = structuredClone(source)
  for (const block of derived.body.blocks) for (const run of block.paragraph?.runs ?? []) if (run.page_field) { run.text = values[run.id]!; run.layout_page_field = run.page_field; delete run.page_field }
  return derived
}

/** Final values derive only from the actual page carrying every field glyph. */
export function nativeDocxBodyPageFieldValuesV1(source: NativeDocxDocumentV1, request: NativeDocxPaginationRequestV1, layout: NativeDocxPaginatedLayoutV1): NativeDocxBodyPageFieldValuesV1 {
  const fields = nativeDocxBodyPageFieldRunsV1(source)
  if (layout.status !== 'paginated' || layout.pages.length > 64) throw new TypeError('Body page fields require bounded successful pagination')
  const paragraphs = new Map(request.shaped_lines.paragraphs.map(p => [p.paragraph_id,p]))
  const locations = new Map(fields.map(run => [run.id,new Set<number>()]))
  for (const page of layout.pages) for (const placed of page.lines) {
    const line = paragraphs.get(placed.paragraph_id)?.lines[placed.source_line_ordinal]
    for (const fragment of line?.fragments ?? []) if (fragment.text.length) locations.get(fragment.source_id)?.add(page.ordinal)
  }
  return Object.fromEntries(fields.map(run => {
    const pages = locations.get(run.id)!
    if (pages.size !== 1) throw new TypeError('Each visible body field must exact-join one final page; hidden or page-split fields refuse')
    return [run.id,String(run.page_field === 'NUMPAGES' ? layout.pages.length : nativeDocxPageNumberV1(source,layout,[...pages][0]!))]
  }))
}

export function validateNativeDocxBodyPageFieldSourceV1(value: unknown, request: NativeDocxPaginationRequestV1, layout: NativeDocxPaginatedLayoutV1): NativeDocxDocumentV1 | undefined {
  if (value === undefined) {
    nativeDocxBodyPageFieldRunsV1(request.document)
    return undefined
  }
  const source = decodeNativeDocxDocument(value)
  if (!source.ok || nativeDocxBodyPageFieldRunsV1(source.value).length === 0) throw new TypeError('Body field source must be one valid original native document with body fields')
  const derived = nativeDocxBodyPageFieldDocumentV1(source.value,nativeDocxBodyPageFieldValuesV1(source.value,request,layout))
  if (JSON.stringify(derived) !== JSON.stringify(request.document)) throw new TypeError('Body field text and all preserved source properties must exactly match final pagination-derived substitution')
  return source.value
}
