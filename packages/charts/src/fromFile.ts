import type { CellRangeRef, ChartSpec, ChartType, NativeChartIdentity } from './types'

// Bridge from the FILE side: the gateway's xlsxpatch reader (Go) summarizes a
// workbook's DrawingML charts as ChartInfo JSON; this module turns that into
// mountable ChartSpecs. It is the "agent made a chart with openpyxl → the
// browser renders it in our layer" path.
//
// Pure functions, tested without Univer: reference parsing (A1 forms, quoted
// sheet names) and spec conversion are exactly the kind of logic that must
// not live inline in a component.

/** Mirror of Go xlsxpatch.ChartInfo (JSON tags). */
export interface FileChartInfo {
  part: string
  identity?: NativeChartIdentity
  type: string
  title?: string
  stacked?: boolean
  series: {
    name?: string
    nameRef?: string
    categoriesRef?: string
    valuesRef?: string
  }[]
}

function validIdentity(identity: NativeChartIdentity | undefined, part: string): identity is NativeChartIdentity {
  return !!identity
    && identity.part === part
    && /^xl\/charts\/[^/]+\.xml$/.test(identity.part)
    && /^xl\/drawings\/(?!_rels\/)[^/]+\.xml$/.test(identity.drawingPart)
    && Number.isSafeInteger(identity.objectId)
    && identity.objectId > 0
    && identity.objectId <= 0xffff_ffff
}

/** Mirror of Go xlsxpatch.ChartAnchor. */
export interface FileChartAnchor {
  FromCol: number
  FromRow: number
  ToCol: number
  ToRow: number
}

const RENDERABLE: ReadonlySet<string> = new Set([
  'column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter', 'radar',
])

/** Parsed "Sheet!$A$1:$B$5" reference: sheet name + 0-indexed inclusive block. */
export interface ParsedRef {
  sheetName: string
  startRow: number
  startColumn: number
  endRow: number
  endColumn: number
}

/** colToIndex("A") = 0, ("Z") = 25, ("AA") = 26 ... */
export function colToIndex(col: string): number {
  let n = 0
  for (const ch of col.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n - 1
}

const CELL_RE = /^\$?([A-Za-z]{1,3})\$?(\d+)$/

/** parseRef handles "Data!$B$2:$B$5", "'My Sheet'!A1", and doubled-quote
 *  escapes. Returns null on anything else — callers skip, never guess. */
export function parseRef(ref: string): ParsedRef | null {
  const bang = ref.lastIndexOf('!')
  if (bang <= 0) return null
  let sheetName = ref.slice(0, bang)
  if (sheetName.startsWith("'") && sheetName.endsWith("'") && sheetName.length >= 2) {
    sheetName = sheetName.slice(1, -1).replace(/''/g, "'")
  }
  if (!sheetName) return null
  const cells = ref.slice(bang + 1).split(':')
  if (cells.length < 1 || cells.length > 2) return null
  const first = CELL_RE.exec(cells[0])
  if (!first) return null
  const last = cells.length === 2 ? CELL_RE.exec(cells[1]) : first
  if (!last) return null
  const startColumn = colToIndex(first[1])
  const startRow = Number(first[2]) - 1
  const endColumn = colToIndex(last[1])
  const endRow = Number(last[2]) - 1
  if (endRow < startRow || endColumn < startColumn) return null
  return { sheetName, startRow, startColumn, endRow, endColumn }
}

/** union merges parsed refs into one bounding block (same sheet only). */
function union(refs: ParsedRef[]): ParsedRef | null {
  if (refs.length === 0) return null
  const sheet = refs[0].sheetName
  let out = { ...refs[0] }
  for (const r of refs.slice(1)) {
    if (r.sheetName !== sheet) return null
    out = {
      sheetName: sheet,
      startRow: Math.min(out.startRow, r.startRow),
      startColumn: Math.min(out.startColumn, r.startColumn),
      endRow: Math.max(out.endRow, r.endRow),
      endColumn: Math.max(out.endColumn, r.endColumn),
    }
  }
  return out
}

export interface FileChartConversion {
  spec: ChartSpec
  /** Grid placement from the file's drawing anchor, when it had one. */
  cellAnchor?: FileChartAnchor
}

/**
 * specFromFileChart converts one file chart into a mountable ChartSpec.
 * Returns null when the chart can't be represented (unrenderable type, no
 * parsable refs, sheet name not in the workbook) — the caller counts those
 * and can say "1 chart in this file isn't shown".
 *
 * The source range is the BOUNDING BLOCK of the series' name/category/value
 * refs including the header/category cells, re-extracted by our heuristics —
 * matching how the chart was almost certainly built from a contiguous table.
 */
export function specFromFileChart(
  info: FileChartInfo,
  sheetIdByName: Record<string, string>,
  anchor?: FileChartAnchor,
): FileChartConversion | null {
  if (!RENDERABLE.has(info.type)) return null
  const refs: ParsedRef[] = []
  for (const s of info.series) {
    for (const r of [s.nameRef, s.categoriesRef, s.valuesRef]) {
      if (!r) continue
      const parsed = parseRef(r)
      if (parsed) refs.push(parsed)
    }
  }
  const block = union(refs)
  if (!block) return null
  const sheetId = sheetIdByName[block.sheetName]
  if (!sheetId) return null
  const range: CellRangeRef = {
    sheetId,
    startRow: block.startRow,
    startColumn: block.startColumn,
    endRow: block.endRow,
    endColumn: block.endColumn,
  }
  const spec: ChartSpec = {
    id: `filechart-${info.part.replace(/[^a-zA-Z0-9]+/g, '-')}`,
    ...(validIdentity(info.identity, info.part) ? { nativeIdentity: { ...info.identity } } : {}),
    type: info.type as ChartType,
    title: info.title || undefined,
    range,
  }
  return { spec, cellAnchor: anchor }
}

/** Convert a whole file's charts; unrepresentable ones are counted, not lost
 *  silently. */
export function specsFromFileCharts(
  charts: FileChartInfo[],
  anchors: Record<string, FileChartAnchor>,
  sheetIdByName: Record<string, string>,
): { conversions: FileChartConversion[]; skipped: number } {
  const conversions: FileChartConversion[] = []
  let skipped = 0
  const identities = new Set<string>()
  for (const c of charts) {
    const identityKey = validIdentity(c.identity, c.part)
      ? `${c.identity.part}\u0000${c.identity.drawingPart}\u0000${c.identity.objectId}`
      : ''
    if (identityKey && identities.has(identityKey)) {
      skipped++
      continue
    }
    const conv = specFromFileChart(c, sheetIdByName, anchors[c.part])
    if (conv) {
      if (identityKey) identities.add(identityKey)
      conversions.push(conv)
    }
    else skipped++
  }
  return { conversions, skipped }
}
