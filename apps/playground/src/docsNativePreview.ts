import {
  decodeNativeDocxDocument,
  type NativeDocxBlockV1,
  type NativeDocxDocumentV1,
  type NativeDocxParagraphV1,
  type NativeDocxRunV1,
  type NativeDocxStoryV1,
} from '../../../packages/docs/src/nativeContract'

export const DOCX_MEDIA_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
export const DOCX_PREVIEW_BLOCK_LIMIT = 200

export type NativeDocxPreviewStats = {
  blocks: number
  paragraphs: number
  tables: number
  textRuns: number
  drawings: number
  references: number
  readOnlyBlocks: number
}

export function decodeNativeDocx(value: unknown): NativeDocxDocumentV1 {
  const decoded = decodeNativeDocxDocument(value)
  if (!decoded.ok) {
    throw new Error(decoded.issues.slice(0, 4).map((issue) => `${issue.path || '/'}: ${issue.message}`).join('; '))
  }
  return decoded.value
}

export function nativeDocxRunText(run: NativeDocxRunV1): string {
  if (run.kind === 'text') return run.properties?.hidden ? '' : run.text ?? ''
  if (run.kind === 'control') {
    if (run.control === 'tab') return '\t'
    if (run.control === 'line-break' || run.control === 'page-break' || run.control === 'column-break') return '\n'
    if (run.control === 'soft-hyphen') return '\u00ad'
  }
  return ''
}

export function nativeDocxParagraphText(paragraph: NativeDocxParagraphV1): string {
  return paragraph.runs.map(nativeDocxRunText).join('')
}

export function nativeDocxStoryText(story: NativeDocxStoryV1): string {
  return story.blocks.map((block) => {
    if (block.paragraph) return nativeDocxParagraphText(block.paragraph)
    return block.table?.rows.flatMap((row) => row.cells.map((cell) => cell.paragraphs.map(nativeDocxParagraphText).join('\n'))).join('\t') ?? ''
  }).filter(Boolean).join('\n')
}

function visitParagraph(paragraph: NativeDocxParagraphV1, stats: NativeDocxPreviewStats): void {
  stats.paragraphs += 1
  if (paragraph.edit_policy.mode === 'read-only') stats.readOnlyBlocks += 1
  for (const run of paragraph.runs) {
    if (run.kind === 'text') stats.textRuns += 1
    if (run.kind === 'drawing') stats.drawings += 1
    if (run.kind === 'reference') stats.references += 1
  }
}

function visitBlock(block: NativeDocxBlockV1, stats: NativeDocxPreviewStats): void {
  stats.blocks += 1
  if (block.paragraph) visitParagraph(block.paragraph, stats)
  if (block.table) {
    stats.tables += 1
    if (block.table.edit_policy.mode === 'read-only') stats.readOnlyBlocks += 1
    for (const row of block.table.rows) for (const cell of row.cells) for (const paragraph of cell.paragraphs) visitParagraph(paragraph, stats)
  }
}

export function nativeDocxPreviewStats(document: NativeDocxDocumentV1): NativeDocxPreviewStats {
  const stats: NativeDocxPreviewStats = { blocks: 0, paragraphs: 0, tables: 0, textRuns: 0, drawings: 0, references: 0, readOnlyBlocks: 0 }
  for (const block of document.body.blocks) visitBlock(block, stats)
  return stats
}

export function visibleNativeDocxBlocks(document: NativeDocxDocumentV1): { blocks: NativeDocxBlockV1[]; omitted: number } {
  const blocks = document.body.blocks.slice(0, DOCX_PREVIEW_BLOCK_LIMIT)
  return { blocks, omitted: document.body.blocks.length - blocks.length }
}
