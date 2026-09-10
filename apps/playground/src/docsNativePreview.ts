import {
  decodeNativeDocxDocument,
  type NativeDocxBlockV1,
  type NativeDocxDocumentV1,
  type NativeDocxParagraphV1,
  type NativeDocxRunV1,
  type NativeDocxStoryV1,
  type NativeDocxTableV1,
  type NativeDocxTableCellV1,
} from '../../../packages/docs/src/nativeContract'

export const DOCX_MEDIA_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
export const DOCX_PREVIEW_BLOCK_LIMIT = 200

export type PreviewTableCell = { cell: NativeDocxTableCellV1; column: number; rowSpan: number; orphanContinuation: boolean }

/** HTML projection only. Merge only identical grid intervals; never discard
 * unexpected continuation content, or join through an intervening row. */
export function nativeDocxTableRows(table: NativeDocxTableV1): PreviewTableCell[][] {
  let active = new Map<string, PreviewTableCell>()
  return table.rows.map((row) => {
    const next = new Map<string, PreviewTableCell>()
    const visible: PreviewTableCell[] = []
    let column = 0
    for (const cell of row.cells) {
      const key = `${column}:${cell.grid_span}`
      const origin = active.get(key)
      const hasContent = cell.paragraphs.some((paragraph) => paragraph.runs.some((run) =>
        run.kind === 'drawing' || run.kind === 'reference' || nativeDocxRunText(run).trim().length > 0))
      if (cell.vertical_merge === 'continue' && origin && !hasContent) {
        origin.rowSpan += 1
        next.set(key, origin)
      } else {
        const projected = { cell, column, rowSpan: 1, orphanContinuation: cell.vertical_merge === 'continue' }
        visible.push(projected)
        if (cell.vertical_merge === 'restart') next.set(key, projected)
      }
      column += cell.grid_span
    }
    active = next
    return visible
  })
}

/** OOXML highlight names aren't CSS colors (notably darkYellow). */
export function nativeDocxHighlight(value: string | undefined): string | undefined {
  const colors: Record<string, string> = {
    black: '#000000', blue: '#0000ff', cyan: '#00ffff', green: '#00ff00',
    magenta: '#ff00ff', red: '#ff0000', yellow: '#ffff00', white: '#ffffff',
    darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000',
    darkMagenta: '#800080', darkRed: '#800000', darkYellow: '#808000',
    darkGray: '#808080', lightGray: '#c0c0c0',
  }
  return value && Object.hasOwn(colors, value) ? colors[value] : undefined
}

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
