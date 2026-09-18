import type { NativeDocxDocumentV1 } from './nativeContract.js'

/** The label alphabet one note kind paints, indexed by counter value.
 *
 * Which counter value a note gets is this tier's decision: the shaper assigns
 * it in body reference order and the paginator and painter re-derive the same
 * assignment from the same source. How that value is *spelled* is the source
 * package's decision, stated by w:footnotePr/w:endnotePr w:numFmt (ECMA-376
 * 17.11.17/17.11.18) and formatted by the extractor with the one counter
 * implementation this codebase has (nativeFormatNumberingCounter). Nothing
 * here re-implements a counter; it only indexes the table the extractor sent.
 *
 * A document that states no modeled w:numFmt carries no table, and every tier
 * keeps painting the decimal counter it painted before. */
export type NativeDocxNoteLabelsV1 = ReadonlyMap<string, readonly string[]>

export function nativeDocxNoteLabelsV1(document: Pick<NativeDocxDocumentV1, 'note_numbering'>): NativeDocxNoteLabelsV1 {
  return new Map((document.note_numbering ?? []).map((entry) => [entry.kind, entry.labels]))
}

/** The label for a 1-based counter value, or undefined when the document
 * states an alphabet that does not reach it. A caller that gets undefined
 * must refuse rather than fall back to the decimal spelling, because the
 * package asked for a different one. */
export function nativeDocxNoteLabelV1(labels: NativeDocxNoteLabelsV1, kind: string, value: number): string | undefined {
  const table = labels.get(kind)
  if (!table) return String(value)
  return table[value - 1]
}
