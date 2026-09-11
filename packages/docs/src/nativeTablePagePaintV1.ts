/**
 * Pure qualification and geometry for the bounded native WordprocessingML
 * table page-paint slice. This module is the only authority for table math;
 * it has no web-layout or document-conversion dependency.
 */

import { sha256 } from '@noble/hashes/sha2.js'
import { bytesToHex } from '@noble/hashes/utils.js'
import type {
  NativeDocxDocumentV1,
  NativeDocxTableBorderV1,
  NativeDocxTableBordersV1,
  NativeDocxTableCellV1,
  NativeDocxTableV1,
} from './nativeContract.js'
import type { NativeDocxResolvedLayoutInputV1 } from './nativeResolvedLayout.js'
import type { NativeDocxShapedLinesV1, NativeDocxShapedParagraphV1 } from './nativeShapingLines.js'

export const DOCX_TABLE_PAGE_PAINT_LIMITS = {
  maxTables: 1_000,
  maxRows: 10_000,
  maxCells: 100_000,
  maxGridColumns: 256,
  maxCoordinateMilliPoints: 1_000_000_000_000,
} as const

export interface NativeDocxQualifiedTableCellV1 {
  cell: NativeDocxTableCellV1
  column_ordinal: number
  grid_span: number
  vertical_merge: NativeDocxTableCellV1['vertical_merge']
  x_millipoints: number
  width_millipoints: number
  content_x_millipoints: number
  content_width_millipoints: number
}

export interface NativeDocxQualifiedTableRowV1 {
  row_id: string
  row_ordinal: number
  cells: NativeDocxQualifiedTableCellV1[]
}

export interface NativeDocxQualifiedTableV1 {
  table: NativeDocxTableV1
  width_millipoints: number
  x_millipoints: number
  grid_widths_millipoints: number[]
  rows: NativeDocxQualifiedTableRowV1[]
}

export interface NativeDocxTableQualificationDiagnosticV1 {
  code: 'unsupported-table-source' | 'table-resource-limit'
  scope_id: string
  message: string
}

export type NativeDocxQualifiedTablesV1 =
  | { status: 'qualified'; tables: NativeDocxQualifiedTableV1[]; paragraph_widths: ReadonlyMap<string, number>; sha256: string }
  | { status: 'refused'; tables: []; paragraph_widths: ReadonlyMap<string, number>; diagnostics: NativeDocxTableQualificationDiagnosticV1[] }

export interface NativeDocxTableCellGeometryV1 {
  cell_id: string
  column_ordinal: number
  grid_span: number
  row_span: number
  vertical_merge: NativeDocxTableCellV1['vertical_merge']
  x_millipoints: number
  y_millipoints: number
  width_millipoints: number
  height_millipoints: number
  content_x_millipoints: number
  content_y_millipoints: number
  content_width_millipoints: number
  content_height_millipoints: number
  shading_rgb?: string
}

export interface NativeDocxTableRowGeometryV1 {
  row_id: string
  row_ordinal: number
  height_millipoints: number
  cells: NativeDocxTableCellGeometryV1[]
}

const MAX_SAFE_TWIPS = Math.floor(DOCX_TABLE_PAGE_PAINT_LIMITS.maxCoordinateMilliPoints / 50)

function checked(...values: number[]): number | undefined {
  let result = 0
  for (const value of values) {
    result += value
    if (!Number.isSafeInteger(result) || Math.abs(result) > DOCX_TABLE_PAGE_PAINT_LIMITS.maxCoordinateMilliPoints) return undefined
  }
  return result
}

function twips(value: number): number | undefined {
  return Number.isSafeInteger(value) && value >= 0 && value <= MAX_SAFE_TWIPS ? checked(value * 50) : undefined
}

function canonical(value: unknown, active = new WeakSet<object>(), depth = 0, state = { nodes: 0 }): unknown {
  state.nodes += 1
  if (state.nodes > 1_000_000 || depth > 64) throw new RangeError('table hash input exceeds the bounded canonicalization budget')
  if (typeof value === 'number' && (!Number.isFinite(value) || Object.is(value, -0))) throw new TypeError('table hash input contains a non-canonical number')
  if (value === null || typeof value !== 'object') return value
  if (active.has(value)) throw new TypeError('table hash input must be acyclic JSON wire data')
  active.add(value)
  const output = Array.isArray(value)
    ? value.map((entry) => canonical(entry, active, depth + 1, state))
    : Object.fromEntries(Object.keys(value).sort().map((key) => [key, canonical((value as Record<string, unknown>)[key], active, depth + 1, state)]))
  active.delete(value)
  return output
}

function projection(tables: readonly NativeDocxQualifiedTableV1[]): unknown {
  return tables.map((entry) => ({
    table_id: entry.table.id,
    width_twips: entry.table.width_twips,
    layout: entry.table.layout,
    alignment: entry.table.alignment,
    indent_twips: entry.table.indent_twips,
    grid_widths_twips: entry.table.grid_widths_twips,
    cell_margins: entry.table.cell_margins,
    borders: entry.table.borders,
    rows: entry.table.rows.map((row) => ({
      row_id: row.id, height_twips: row.height_twips, height_rule: row.height_rule, repeat_header: row.repeat_header, cant_split: row.cant_split,
      cells: row.cells.map((cell) => ({
        cell_id: cell.id, width_twips: cell.width_twips, grid_span: cell.grid_span, vertical_merge: cell.vertical_merge,
        borders: cell.borders, shading_rgb: cell.shading_rgb, paragraph_ids: cell.paragraphs.map((paragraph) => paragraph.id),
      })),
    })),
  }))
}

/** Hashes the exact already-qualified table projection; object-key order is irrelevant and source-array order is authoritative. */
export function nativeDocxTableProjectionSha256V1(tables: readonly NativeDocxQualifiedTableV1[]): string {
  try {
    if (!Array.isArray(tables) || tables.length > DOCX_TABLE_PAGE_PAINT_LIMITS.maxTables) throw new TypeError('table hash input must be a bounded array')
    let rows = 0
    let cells = 0
    for (const table of tables) {
      if (!table || !Array.isArray(table.table?.rows) || !Array.isArray(table.table.grid_widths_twips) || table.table.grid_widths_twips.length > DOCX_TABLE_PAGE_PAINT_LIMITS.maxGridColumns) throw new TypeError('table hash input is not a qualified projection')
      rows += table.table.rows.length
      if (rows > DOCX_TABLE_PAGE_PAINT_LIMITS.maxRows) throw new RangeError('table hash rows exceed the bounded canonicalization budget')
      for (const row of table.table.rows) {
        if (!Array.isArray(row.cells)) throw new TypeError('table hash row cells are malformed')
        cells += row.cells.length
        if (cells > DOCX_TABLE_PAGE_PAINT_LIMITS.maxCells) throw new RangeError('table hash cells exceed the bounded canonicalization budget')
      }
    }
    return `sha256:${bytesToHex(sha256(new TextEncoder().encode(JSON.stringify(canonical(projection(tables))))))}`
  } catch (error) {
    if (error instanceof RangeError) throw error
    throw new TypeError('table hash input must be bounded cloneable canonical wire data')
  }
}

function borderValid(border: NativeDocxTableBorderV1 | undefined): boolean {
  if (!border) return true
  if (border.style === 'none') return border.size_eighth_points === 0 && border.color_rgb === undefined
  return border.style === 'single' && Number.isSafeInteger(border.size_eighth_points) && border.size_eighth_points > 0 && border.size_eighth_points <= 768 && /^[0-9A-F]{6}$/.test(border.color_rgb ?? '')
}

function bordersValid(borders: NativeDocxTableBordersV1 | undefined): boolean {
  return !borders || Object.values(borders).every(borderValid)
}

function cellRunsSupported(cell: NativeDocxTableCellV1): boolean {
  return cell.paragraphs.length > 0 && cell.paragraphs.every((paragraph) => paragraph.runs.every((run) => run.kind === 'text' || run.kind === 'control' && (run.control === 'tab' || run.control === 'line-break')))
}

function cellHasVisibleContent(cell: NativeDocxTableCellV1): boolean {
  return cell.paragraphs.some((paragraph) => paragraph.runs.some((run) => (run.kind === 'text' && (run.text ?? '') !== '') || run.kind === 'drawing' || (run.kind === 'control' && run.control !== 'tab' && run.control !== 'line-break')))
}

function tableStyleBlocksPaint(resolved: NativeDocxResolvedLayoutInputV1, tableID: string): boolean {
  return resolved.diagnostics.some((diagnostic) => diagnostic.scope_id === tableID && (diagnostic.code === 'CONDITIONAL_TABLE_STYLE_PRESERVED' || diagnostic.code === 'TABLE_STYLE_EFFECTS_PRESERVED' || diagnostic.code === 'MISSING_TABLE_STYLE'))
}

/**
 * Qualifies a deliberately narrow exact subset. Any ambiguity refuses the
 * entire table set before shaping or outline-provider work begins.
 */
export function qualifyNativeDocxTablesV1(document: NativeDocxDocumentV1, resolved: NativeDocxResolvedLayoutInputV1): NativeDocxQualifiedTablesV1 {
  const sourceTables = document.body.blocks.flatMap((block) => block.kind === 'table' && block.table ? [block.table] : [])
  if (sourceTables.length === 0) return { status: 'qualified', tables: [], paragraph_widths: new Map(), sha256: nativeDocxTableProjectionSha256V1([]) }
  const fail = (scope_id: string, message: string): NativeDocxQualifiedTablesV1 => ({ status: 'refused', tables: [], paragraph_widths: new Map(), diagnostics: [{ code: 'unsupported-table-source', scope_id, message }] })
  if (sourceTables.length > DOCX_TABLE_PAGE_PAINT_LIMITS.maxTables) return { status: 'refused', tables: [], paragraph_widths: new Map(), diagnostics: [{ code: 'table-resource-limit', scope_id: document.document_id, message: `Tables exceed ${DOCX_TABLE_PAGE_PAINT_LIMITS.maxTables}` }] }
  const resolvedParagraphs = new Map(resolved.paragraphs.map((entry) => [entry.paragraph_id, entry]))
  const resolvedTables = new Map(resolved.tables.map((entry) => [entry.table_id, entry]))
  const tables: NativeDocxQualifiedTableV1[] = []
  const paragraphWidths = new Map<string, number>()
  let rows = 0
  let cells = 0
  for (const [blockIndex, block] of document.body.blocks.entries()) if (block.table) {
    const previous = document.body.blocks[blockIndex - 1]?.paragraph
    const next = document.body.blocks[blockIndex + 1]?.paragraph
    if (previous && (resolvedParagraphs.get(previous.id)?.properties.spacing_after_twips ?? 0) !== 0) return fail(block.table.id, 'Paragraph spacing adjacent to a table must be explicit zero in v1')
    if (next && (resolvedParagraphs.get(next.id)?.properties.spacing_before_twips ?? 0) !== 0) return fail(block.table.id, 'Paragraph spacing adjacent to a table must be explicit zero in v1')
  }
  for (const sourceTable of sourceTables) {
    const resolvedTable = resolvedTables.get(sourceTable.id)
    if (!resolvedTable) return fail(sourceTable.id, 'Table does not exact-join the resolved layout')
    if (tableStyleBlocksPaint(resolved, sourceTable.id)) return fail(sourceTable.id, 'Table styles and conditional style effects are outside the bounded page-paint subset')
    const paintBorders = sourceTable.borders ?? resolvedTable.borders
    if ((sourceTable.table_style_id || resolvedTable.style_id) && !bordersValid(paintBorders)) return fail(sourceTable.id, 'Simple table style did not project exact table-level border commands')
    const table = paintBorders === sourceTable.borders ? sourceTable : { ...sourceTable, borders: paintBorders }
    if (table.layout !== 'fixed' || table.alignment !== 'left' || table.indent_twips === undefined || table.width_twips === undefined || !table.cell_margins) return fail(table.id, 'Table requires explicit fixed dxa width, left alignment, indent, and all four cell margins')
    const grid = table.grid_widths_twips
    if (!grid || grid.length === 0 || grid.length > DOCX_TABLE_PAGE_PAINT_LIMITS.maxGridColumns || grid.some((width) => !Number.isSafeInteger(width) || width <= 0 || width > MAX_SAFE_TWIPS)) return fail(table.id, 'Table requires a bounded non-empty positive tblGrid')
    const gridSum = grid.reduce((sum, width) => sum + width, 0)
    if (!Number.isSafeInteger(gridSum) || gridSum !== table.width_twips) return fail(table.id, 'Table width must exactly equal the sum of fixed grid columns')
    if (!bordersValid(table.borders)) return fail(table.id, 'Table border style, width, or RGB color is outside the bounded subset')
    const gridMP = grid.map((width) => twips(width)!)
    const tableWidth = twips(table.width_twips)
    const tableX = twips(table.indent_twips)
    const leftMargin = twips(table.cell_margins.left_twips)
    const rightMargin = twips(table.cell_margins.right_twips)
    if (tableWidth === undefined || tableX === undefined || checked(tableX, tableWidth) === undefined || leftMargin === undefined || rightMargin === undefined) return fail(table.id, 'Table geometry exceeds bounded integer milli-points')
    if (table.rows.length === 0) return fail(table.id, 'Empty tables are outside the bounded page-paint subset')
    const qualifiedRows: NativeDocxQualifiedTableRowV1[] = []
    rows += table.rows.length
    if (rows > DOCX_TABLE_PAGE_PAINT_LIMITS.maxRows) return { status: 'refused', tables: [], paragraph_widths: new Map(), diagnostics: [{ code: 'table-resource-limit', scope_id: table.id, message: `Rows exceed ${DOCX_TABLE_PAGE_PAINT_LIMITS.maxRows}` }] }
    const openMerge: Array<{ cellID: string; span: number } | undefined> = Array.from({ length: grid.length })
    let bodyStarted = false
    const repeating = table.rows.some((row) => row.repeat_header)
    for (const [rowOrdinal, row] of table.rows.entries()) {
      if (row.repeat_header && bodyStarted) return fail(row.id, 'Repeated headers must be a contiguous leading row prefix')
      if (!row.repeat_header) bodyStarted = true
      if (repeating && row.cells.some((cell) => cell.vertical_merge !== 'none')) return fail(row.id, 'Vertical merges in repeating-header tables require a separate pagination contract')
      if (row.cant_split !== true) return fail(row.id, 'Every qualified row must explicitly prohibit page splitting')
      if ((row.height_twips === undefined) !== (row.height_rule === undefined)) return fail(row.id, 'Row height requires both height_twips and height_rule')
      if (row.height_twips !== undefined && (row.height_rule !== 'atLeast' && row.height_rule !== 'exact' || !Number.isSafeInteger(row.height_twips) || row.height_twips < 0 || row.height_twips > MAX_SAFE_TWIPS)) return fail(row.id, 'Row height rule is outside the exact atLeast/exact subset')
      cells += row.cells.length
      if (cells > DOCX_TABLE_PAGE_PAINT_LIMITS.maxCells) return { status: 'refused', tables: [], paragraph_widths: new Map(), diagnostics: [{ code: 'table-resource-limit', scope_id: table.id, message: `Cells exceed ${DOCX_TABLE_PAGE_PAINT_LIMITS.maxCells}` }] }
      let x = tableX
      let column = 0
      const qualifiedCells: NativeDocxQualifiedTableCellV1[] = []
      const consumed = new Array<boolean>(grid.length).fill(false)
      for (const cell of row.cells) {
        const span = cell.grid_span
        if (!Number.isSafeInteger(span) || span < 1 || column + span > grid.length) return fail(cell.id, 'Horizontal merge must consume a contiguous in-range tblGrid slice')
        let widthTwips = 0
        const widthParts: number[] = []
        for (let offset = 0; offset < span; offset += 1) {
          if (consumed[column + offset]) return fail(cell.id, 'Merged cells overlap on the fixed tblGrid')
          consumed[column + offset] = true
          widthTwips += grid[column + offset]!
          widthParts.push(gridMP[column + offset]!)
        }
        if (cell.width_twips !== widthTwips) return fail(cell.id, 'Cell width must exactly equal the sum of its fixed grid columns')
        if (cell.borders) return fail(cell.id, 'Cell border conflict resolution is outside v1; use unambiguous table-level borders')
        const shading = cell.shading_rgb ?? resolvedTable.cell_shading_rgb
        if (shading !== undefined && !/^[0-9A-F]{6}$/.test(shading)) return fail(cell.id, 'Cell shading must be an explicit RGB clear fill')
        if (!cellRunsSupported(cell)) return fail(cell.id, 'Cells must contain direct paragraphs with text, tab, or line-break runs only')
        if (cell.vertical_merge === 'continue' && cellHasVisibleContent(cell)) return fail(cell.id, 'Vertical-merge continue cells cannot carry independent visible content')
        for (let offset = 0; offset < span; offset += 1) {
          const open = openMerge[column + offset]
          if (cell.vertical_merge === 'continue') {
            if (!open || open.span !== span || (offset > 0 && openMerge[column] !== open)) return fail(cell.id, 'Vertical-merge continue does not exact-join a restart covering the same grid columns')
          }
        }
        if (cell.vertical_merge !== 'continue') {
          const marker = cell.vertical_merge === 'restart' ? { cellID: cell.id, span } : undefined
          for (let offset = 0; offset < span; offset += 1) openMerge[column + offset] = marker
        }
        const width = span === 1 ? widthParts[0]! : checked(...widthParts)
        const contentX = checked(x, leftMargin)
        const contentWidth = width === undefined ? undefined : checked(width, -leftMargin, -rightMargin)
        if (width === undefined || contentX === undefined || contentWidth === undefined || contentWidth <= 0) return fail(cell.id, 'Cell margins leave a non-positive or unbounded paragraph width')
        const paintCell = shading === cell.shading_rgb ? cell : { ...cell, shading_rgb: shading }
        for (const paragraph of cell.paragraphs) {
          if (!resolvedParagraphs.has(paragraph.id)) return fail(paragraph.id, 'Cell paragraph does not exact-join the resolved layout')
          if (resolvedParagraphs.get(paragraph.id)?.numbering) return fail(paragraph.id, 'Numbering inside tables is refused because list-counter state is not guessed')
          paragraphWidths.set(paragraph.id, contentWidth)
        }
        qualifiedCells.push({ cell: paintCell, column_ordinal: column, grid_span: span, vertical_merge: cell.vertical_merge, x_millipoints: x, width_millipoints: width, content_x_millipoints: contentX, content_width_millipoints: contentWidth })
        const nextX = checked(x, width)
        if (nextX === undefined) return fail(cell.id, 'Cell geometry exceeds bounded integer milli-points')
        x = nextX
        column += span
      }
      if (column !== grid.length || consumed.some((value) => !value)) return fail(row.id, 'Every row must consume the exact fixed grid without omitted cells')
      qualifiedRows.push({ row_id: row.id, row_ordinal: rowOrdinal, cells: qualifiedCells })
    }
    tables.push({ table, width_millipoints: tableWidth, x_millipoints: tableX, grid_widths_millipoints: gridMP, rows: qualifiedRows })
  }
  if (resolved.diagnostics.some((diagnostic) => sourceTables.some((table) => diagnostic.scope_id === table.id || table.rows.some((row) => row.cells.some((cell) => cell.id === diagnostic.scope_id || cell.paragraphs.some((paragraph) => paragraph.id === diagnostic.scope_id || paragraph.runs.some((run) => run.id === diagnostic.scope_id))))))) return fail(document.document_id, 'Resolved-layout diagnostics touch a table or descendant and exact table paint is unavailable')
  return { status: 'qualified', tables, paragraph_widths: paragraphWidths, sha256: nativeDocxTableProjectionSha256V1(tables) }
}

function cellContentHeight(cell: NativeDocxQualifiedTableCellV1, paragraphs: Map<string, NativeDocxShapedParagraphV1>): number | undefined {
  let contentHeight = 0
  let previousAfter = 0
  for (const [paragraphIndex, source] of cell.cell.paragraphs.entries()) {
    const paragraph = paragraphs.get(source.id)
    if (!paragraph) return undefined
    const lines = paragraph.lines.reduce((sum, line) => sum + line.line_height_millipoints, 0)
    const gap = paragraphIndex === 0 ? paragraph.spacing_before_millipoints : Math.max(previousAfter, paragraph.spacing_before_millipoints)
    contentHeight = checked(contentHeight, gap, lines) ?? -1
    previousAfter = paragraph.spacing_after_millipoints
  }
  contentHeight = checked(contentHeight, previousAfter) ?? -1
  return contentHeight < 0 ? undefined : contentHeight
}

function mergeRowSpan(table: NativeDocxQualifiedTableV1, rowIndex: number, cell: NativeDocxQualifiedTableCellV1): number {
  if (cell.vertical_merge !== 'restart') return 1
  let span = 1
  while (rowIndex + span < table.rows.length) {
    const below = table.rows[rowIndex + span]!.cells.find((candidate) => candidate.column_ordinal === cell.column_ordinal && candidate.grid_span === cell.grid_span && candidate.vertical_merge === 'continue')
    if (!below) break
    span += 1
  }
  return span
}

export function nativeDocxTableRowGroupSizeV1(table: NativeDocxQualifiedTableV1, rowIndex: number): number {
  let span = 1
  for (const cell of table.rows[rowIndex]?.cells ?? []) span = Math.max(span, mergeRowSpan(table, rowIndex, cell))
  return span
}

/** Derives deterministic row/cell boxes before page placement. */
export function layoutNativeDocxTableRowsV1(table: NativeDocxQualifiedTableV1, shaped: NativeDocxShapedLinesV1): NativeDocxTableRowGeometryV1[] | undefined {
  const paragraphs = new Map(shaped.paragraphs.map((entry) => [entry.paragraph_id, entry]))
  const margins = table.table.cell_margins!
  const top = twips(margins.top_twips)!
  const right = twips(margins.right_twips)!
  const bottom = twips(margins.bottom_twips)!
  const left = twips(margins.left_twips)!
  const output: NativeDocxTableRowGeometryV1[] = []
  for (const row of table.rows) {
    const sourceRow = table.table.rows[row.row_ordinal]!
    const contentHeights: number[] = []
    const owners = row.cells.filter((cell) => cell.vertical_merge !== 'continue')
    const heightSource = owners.length > 0 ? owners : row.cells
    let rowHeight = 0
    for (const cell of row.cells) {
      const contentHeight = cellContentHeight(cell, paragraphs)
      if (contentHeight === undefined) return undefined
      contentHeights.push(contentHeight)
    }
    for (const cell of heightSource) {
      const contentHeight = contentHeights[row.cells.indexOf(cell)]!
      rowHeight = Math.max(rowHeight, checked(top, contentHeight, bottom) ?? Number.MAX_SAFE_INTEGER)
    }
    if (sourceRow.height_rule === 'atLeast' && sourceRow.height_twips !== undefined) {
      const minimum = twips(sourceRow.height_twips)
      if (minimum === undefined) return undefined
      rowHeight = Math.max(rowHeight, minimum)
    } else if (sourceRow.height_rule === 'exact' && sourceRow.height_twips !== undefined) {
      const exactHeight = twips(sourceRow.height_twips)
      if (exactHeight === undefined || exactHeight <= 0) return undefined
      if (rowHeight > exactHeight) return undefined
      rowHeight = exactHeight
    }
    if (!Number.isSafeInteger(rowHeight) || rowHeight <= 0 || rowHeight > DOCX_TABLE_PAGE_PAINT_LIMITS.maxCoordinateMilliPoints) return undefined
    output.push({
      row_id: row.row_id,
      row_ordinal: row.row_ordinal,
      height_millipoints: rowHeight,
      cells: row.cells.map((cell, index) => ({
        cell_id: cell.cell.id, column_ordinal: cell.column_ordinal, grid_span: cell.grid_span, row_span: 1, vertical_merge: cell.vertical_merge,
        x_millipoints: cell.x_millipoints, y_millipoints: 0,
        width_millipoints: cell.width_millipoints, height_millipoints: rowHeight,
        content_x_millipoints: checked(cell.x_millipoints, left)!, content_y_millipoints: top,
        content_width_millipoints: checked(cell.width_millipoints, -left, -right)!,
        content_height_millipoints: contentHeights[index]!,
        ...(cell.cell.shading_rgb ? { shading_rgb: cell.cell.shading_rgb } : {}),
      })),
    })
  }
  for (const [rowIndex, row] of table.rows.entries()) {
    for (const [cellIndex, cell] of row.cells.entries()) {
      if (cell.vertical_merge !== 'restart') continue
      const span = mergeRowSpan(table, rowIndex, cell)
      let height = 0
      for (let offset = 0; offset < span; offset += 1) height = checked(height, output[rowIndex + offset]!.height_millipoints) ?? -1
      if (height < 0) return undefined
      const contentBox = checked(height, -top, -bottom)
      if (contentBox === undefined || output[rowIndex]!.cells[cellIndex]!.content_height_millipoints > contentBox) return undefined
      output[rowIndex]!.cells[cellIndex]!.row_span = span
      output[rowIndex]!.cells[cellIndex]!.height_millipoints = height
    }
  }
  return output
}
