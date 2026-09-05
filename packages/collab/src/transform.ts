// Stage C — structural transforms.
//
// The gateway keeps the op log LINEAR (a submission whose base isn't the
// head is rejected with STALE_BASE; the client catches up and resubmits),
// so history never needs transforming. What remains is the classic
// two-party problem between an INCOMING entry and this client's PENDING
// (unacknowledged) ops: the incoming entry was ordered first, our pending
// ops will be ordered after it, but we already applied ours locally.
//
//   apply locally:  T(incoming, pending)  — shift by our pending structural
//                                           edits; drop cells our pending
//                                           writes own (ours end up later)
//   keep shipping:  T(pending, incoming)  — shift by the incoming entry's
//                                           structural edits before resubmit
//
// Geometry covered: row/col insert, remove, and move plus sheet removal, applied to
// cell matrices (set-range-values), IRange-bearing params (merges, heights,
// the structural ops themselves) and peer selections. Content-only range
// moves/reorders participate in cell ownership; their interaction with later
// structural geometry remains fail-closed because the native mutations do not
// expose a split-range representation.

import type { SheetSelection } from './types'
import { cellsOf, MOVE_RANGE, REORDER_RANGE, SET_RANGE_VALUES, type OpRecord } from './sync'

export type Axis = 'row' | 'col'

/** A structural edit in normalized form. */
export interface StructEdit {
  kind: 'insert' | 'remove'
  axis: Axis
  subUnitId: string
  /** First affected index. */
  start: number
  count: number
}

/** A row/column block permutation. Target coordinates are the final range. */
export interface MoveEdit {
  kind: 'move'
  axis: Axis
  subUnitId: string
  sourceStart: number
  sourceEnd: number
  targetStart: number
  targetEnd: number
}

interface IRangeLike {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
  [k: string]: unknown
}

const INSERT_ROW = 'sheet.mutation.insert-row'
const INSERT_COL = 'sheet.mutation.insert-col'
const REMOVE_ROWS = 'sheet.mutation.remove-rows'
const REMOVE_COL = 'sheet.mutation.remove-col'
const REMOVE_SHEET = 'sheet.mutation.remove-sheet'
const CONTENT_RANGE_OPS = new Set<string>([MOVE_RANGE, REORDER_RANGE])

function isRange(v: unknown): v is IRangeLike {
  if (!v || typeof v !== 'object') return false
  const r = v as Record<string, unknown>
  return (
    typeof r.startRow === 'number' && typeof r.endRow === 'number' && typeof r.startColumn === 'number' && typeof r.endColumn === 'number'
  )
}

/** Normalize a structural mutation; null when it isn't one we model. */
export function parseStruct(op: OpRecord): StructEdit | null {
  const p = op.params as { subUnitId?: unknown; range?: unknown }
  if (typeof p?.subUnitId !== 'string' || !isRange(p.range)) return null
  const r = p.range
  switch (op.id) {
    case INSERT_ROW:
      return { kind: 'insert', axis: 'row', subUnitId: p.subUnitId, start: r.startRow, count: r.endRow - r.startRow + 1 }
    case REMOVE_ROWS:
      return { kind: 'remove', axis: 'row', subUnitId: p.subUnitId, start: r.startRow, count: r.endRow - r.startRow + 1 }
    case INSERT_COL:
      return { kind: 'insert', axis: 'col', subUnitId: p.subUnitId, start: r.startColumn, count: r.endColumn - r.startColumn + 1 }
    case REMOVE_COL:
      return { kind: 'remove', axis: 'col', subUnitId: p.subUnitId, start: r.startColumn, count: r.endColumn - r.startColumn + 1 }
    default:
      return null
  }
}

/** Normalize Univer's row/column move mutations. Invalid/non-equal ranges are refused. */
export function parseMove(op: OpRecord): MoveEdit | null {
  if (op.id !== 'sheet.mutation.move-rows' && op.id !== 'sheet.mutation.move-columns') return null
  const p = op.params as { subUnitId?: unknown; sourceRange?: unknown; targetRange?: unknown; fromRange?: unknown; toRange?: unknown }
  if (typeof p?.subUnitId !== 'string') return null
  const source = isRange(p.sourceRange) ? p.sourceRange : isRange(p.fromRange) ? p.fromRange : null
  const target = isRange(p.targetRange) ? p.targetRange : isRange(p.toRange) ? p.toRange : null
  if (!source || !target) return null
  const axis: Axis = op.id === 'sheet.mutation.move-rows' ? 'row' : 'col'
  const sourceStart = axis === 'row' ? source.startRow : source.startColumn
  const sourceEnd = axis === 'row' ? source.endRow : source.endColumn
  const targetStart = axis === 'row' ? target.startRow : target.startColumn
  const targetEnd = axis === 'row' ? target.endRow : target.endColumn
  if (![sourceStart, sourceEnd, targetStart, targetEnd].every(Number.isSafeInteger)) return null
  if (sourceStart < 0 || targetStart < 0 || sourceEnd < sourceStart || targetEnd < targetStart) return null
  if (sourceEnd - sourceStart !== targetEnd - targetStart) return null
  if (sourceStart !== targetStart && sourceStart <= targetEnd && targetStart <= sourceEnd) return null
  return { kind: 'move', axis, subUnitId: p.subUnitId, sourceStart, sourceEnd, targetStart, targetEnd }
}

/** Map one coordinate through a completed block move. */
export function mapIndexThroughMove(move: MoveEdit, index: number): number {
  const { sourceStart: source, sourceEnd, targetStart: target, targetEnd } = move
  const count = sourceEnd - source + 1
  if (target === source) return index
  if (index >= source && index <= sourceEnd) return target + index - source
  if (target < source && index >= target && index < source) return index + count
  if (target > sourceEnd && index > sourceEnd && index <= targetEnd) return index - count
  return index
}

/** Exact mapped union for an inclusive span, returned as sorted disjoint spans. */
export function mapSpanThroughMove(move: MoveEdit, start: number, end: number): Array<[number, number]> {
  if (start > end) return []
  const cuts = new Set([start, end + 1])
  for (const value of [move.sourceStart, move.sourceEnd + 1, move.targetStart, move.targetEnd + 1]) {
    if (value > start && value <= end) cuts.add(value)
  }
  const points = [...cuts].sort((a, b) => a - b)
  const spans: Array<[number, number]> = []
  for (let i = 0; i < points.length - 1; i++) {
    const left = points[i]!
    const right = points[i + 1]! - 1
    const mapped: [number, number] = [mapIndexThroughMove(move, left), mapIndexThroughMove(move, right)]
    if (mapped[0] > mapped[1]) mapped.reverse()
    spans.push(mapped)
  }
  spans.sort((a, b) => a[0] - b[0])
  const merged: Array<[number, number]> = []
  for (const span of spans) {
    const last = merged.at(-1)
    if (last && span[0] <= last[1] + 1) last[1] = Math.max(last[1], span[1])
    else merged.push([...span])
  }
  return merged
}

/** Map one index through an edit; null = the index was removed. */
export function mapIndex(edit: StructEdit, i: number): number | null {
  if (edit.kind === 'insert') return i >= edit.start ? i + edit.count : i
  const end = edit.start + edit.count
  if (i >= edit.start && i < end) return null
  return i >= end ? i - edit.count : i
}

/**
 * Map an inclusive [start, end] span on the edit's axis; null when the span
 * is entirely removed. Partial removals clamp; an insert strictly inside a
 * span widens it (the span's content grew), and at its start shifts it.
 */
export function mapSpan(edit: StructEdit, start: number, end: number): [number, number] | null {
  if (edit.kind === 'insert') {
    if (edit.start <= start) return [start + edit.count, end + edit.count]
    if (edit.start <= end) return [start, end + edit.count]
    return [start, end]
  }
  const remEnd = edit.start + edit.count - 1
  if (edit.start > end) return [start, end] // removal entirely below/right of the span
  if (edit.start <= start && remEnd >= end) return null
  let s = start
  let e = end
  if (remEnd < start) {
    s -= edit.count
    e -= edit.count
  } else if (edit.start <= start) {
    // removal chops the leading part
    s = edit.start
    e -= edit.count
  } else if (remEnd <= end) {
    // removal strictly inside
    e -= edit.count
  } else {
    // removal chops the trailing part
    e = edit.start - 1
  }
  return s <= e ? [s, e] : null
}

function mapRange(edit: StructEdit, r: IRangeLike): IRangeLike | null {
  const span = edit.axis === 'row' ? mapSpan(edit, r.startRow, r.endRow) : mapSpan(edit, r.startColumn, r.endColumn)
  if (!span) return null
  return edit.axis === 'row'
    ? { ...r, startRow: span[0], endRow: span[1] }
    : { ...r, startColumn: span[0], endColumn: span[1] }
}

function mapRangeThroughMove(move: MoveEdit, range: IRangeLike): IRangeLike[] {
  const spans = move.axis === 'row'
    ? mapSpanThroughMove(move, range.startRow, range.endRow)
    : mapSpanThroughMove(move, range.startColumn, range.endColumn)
  return spans.map((span) => move.axis === 'row'
    ? { ...range, startRow: span[0], endRow: span[1] }
    : { ...range, startColumn: span[0], endColumn: span[1] })
}

type CellMatrix = Record<string, Record<string, unknown>>

function mapCellMatrix(edit: StructEdit, m: CellMatrix): CellMatrix | null {
  const out: CellMatrix = {}
  let n = 0
  for (const rk of Object.keys(m)) {
    const row = m[rk]
    if (!row) continue
    const r = Number(rk)
    const mr = edit.axis === 'row' ? mapIndex(edit, r) : r
    if (mr === null) continue
    for (const ck of Object.keys(row)) {
      const c = Number(ck)
      const mc = edit.axis === 'col' ? mapIndex(edit, c) : c
      if (mc === null) continue
      const target = out[String(mr)] ?? (out[String(mr)] = {})
      target[String(mc)] = row[ck]
      n++
    }
  }
  return n > 0 ? out : null
}

function mapCellMatrixThroughMove(move: MoveEdit, matrix: CellMatrix): CellMatrix | null {
  const out: CellMatrix = {}
  let count = 0
  for (const [rowKey, row] of Object.entries(matrix)) {
    if (!row) continue
    const rowIndex = Number(rowKey)
    if (!Number.isSafeInteger(rowIndex)) continue
    const mappedRow = move.axis === 'row' ? mapIndexThroughMove(move, rowIndex) : rowIndex
    for (const [columnKey, value] of Object.entries(row)) {
      const columnIndex = Number(columnKey)
      if (!Number.isSafeInteger(columnIndex)) continue
      const mappedColumn = move.axis === 'col' ? mapIndexThroughMove(move, columnIndex) : columnIndex
      const target = out[String(mappedRow)] ?? (out[String(mappedRow)] = {})
      target[String(mappedColumn)] = value
      count++
    }
  }
  return count > 0 ? out : null
}

function subUnitOf(op: OpRecord): string | undefined {
  const p = op.params as { subUnitId?: unknown }
  return typeof p.subUnitId === 'string' ? p.subUnitId : undefined
}

/**
 * Transform ONE op through ONE structural edit that is ordered before it.
 * Returns the shifted op, the op unchanged (different sheet / no geometry),
 * or null when the edit consumed it entirely.
 */
export function transformOp(op: OpRecord, edit: StructEdit): OpRecord | null {
  if (subUnitOf(op) !== edit.subUnitId) return op
  // Cell writes: remap the matrix.
  if (op.id === SET_RANGE_VALUES) {
    const p = op.params as { cellValue?: CellMatrix | null }
    if (!p.cellValue) return op
    const mapped = mapCellMatrix(edit, p.cellValue)
    if (!mapped) return null
    return { id: op.id, params: { ...op.params, cellValue: mapped } }
  }
  // Structural vs structural on the same axis: shift the op's own span.
  const struct = parseStruct(op)
  if (struct) {
    if (struct.axis !== edit.axis) return op
    if (struct.kind === 'insert') {
      // An insert is a position, not a span: shift its anchor. Equal anchors
      // keep the earlier (incoming) edit first — mapIndex's >= does exactly
      // that, and both sides compute it the same way.
      const at = mapIndex(edit, struct.start)
      if (at === null) {
        // Inserting inside rows someone just removed: land at the cut.
        return remakeStruct(op, edit.start, edit.start + struct.count - 1)
      }
      return remakeStruct(op, at, at + struct.count - 1)
    }
    const span = mapSpan(edit, struct.start, struct.start + struct.count - 1)
    if (!span) return null
    return remakeStruct(op, span[0], span[1])
  }
  // A queued move can be shifted through an earlier insert/remove when both
  // endpoints remain single, equal-sized ranges. If either endpoint is
  // consumed, the move itself no longer has a safe target.
  const move = parseMove(op)
  if (move) {
    return transformMoveThroughStruct(op, move, edit).op
  }
  // Generic range-bearing params (merges, row heights, …).
  const p = op.params as Record<string, unknown>
  if (isRange(p.range)) {
    const mapped = mapRange(edit, p.range)
    if (!mapped) return null
    return { id: op.id, params: { ...p, range: mapped } }
  }
  if (Array.isArray(p.ranges) && p.ranges.every(isRange)) {
    const mapped = (p.ranges as IRangeLike[]).map((r) => mapRange(edit, r)).filter((r): r is IRangeLike => r !== null)
    if (mapped.length === 0) return null
    return { id: op.id, params: { ...p, ranges: mapped } }
  }
  return op
}

function transformMoveThroughStruct(op: OpRecord, move: MoveEdit, edit: StructEdit): { op: OpRecord | null; lossy: boolean } {
  if (move.axis !== edit.axis) return { op, lossy: false }
  const p = op.params as Record<string, unknown>
  const sourceKey = isRange(p.sourceRange) ? 'sourceRange' : 'fromRange'
  const targetKey = isRange(p.targetRange) ? 'targetRange' : 'toRange'
  const sourceRange = p[sourceKey] as IRangeLike
  const targetRange = p[targetKey] as IRangeLike
  const source = mapRange(edit, sourceRange)
  const target = mapRange(edit, targetRange)
  if (!source || !target) return { op: null, lossy: false }
  const sourceLength = move.axis === 'row' ? source.endRow - source.startRow : source.endColumn - source.startColumn
  const targetLength = move.axis === 'row' ? target.endRow - target.startRow : target.endColumn - target.startColumn
  if (sourceLength !== targetLength) return { op, lossy: true }
  return { op: { id: op.id, params: { ...p, [sourceKey]: source, [targetKey]: target } }, lossy: false }
}

interface MoveOpResult {
  ops: OpRecord[]
  lossy: boolean
}

/** Transform one op through a row/column move ordered before it. */
function transformOpThroughMove(op: OpRecord, move: MoveEdit): MoveOpResult {
  if (subUnitOf(op) !== move.subUnitId) return { ops: [op], lossy: false }
  if (op.id === SET_RANGE_VALUES) {
    const params = op.params as { cellValue?: CellMatrix | null }
    if (!params.cellValue) return { ops: [op], lossy: false }
    const mapped = mapCellMatrixThroughMove(move, params.cellValue)
    return { ops: mapped ? [{ id: op.id, params: { ...op.params, cellValue: mapped } }] : [], lossy: false }
  }
  const struct = parseStruct(op)
  if (struct && struct.axis === move.axis) {
    const spans = mapSpanThroughMove(move, struct.start, struct.start + struct.count - 1)
    if (spans.length !== 1) return { ops: [op], lossy: true }
    return { ops: [remakeStruct(op, spans[0]![0], spans[0]![1])], lossy: false }
  }
  const nestedMove = parseMove(op)
  if (nestedMove && nestedMove.axis === move.axis) {
    const source = mapSpanThroughMove(move, nestedMove.sourceStart, nestedMove.sourceEnd)
    const target = mapSpanThroughMove(move, nestedMove.targetStart, nestedMove.targetEnd)
    if (source.length !== 1 || target.length !== 1) return { ops: [op], lossy: true }
    const p = op.params as Record<string, unknown>
    const sourceKey = isRange(p.sourceRange) ? 'sourceRange' : 'fromRange'
    const targetKey = isRange(p.targetRange) ? 'targetRange' : 'toRange'
    const sourceRange = p[sourceKey] as IRangeLike
    const targetRange = p[targetKey] as IRangeLike
    const mappedSource = move.axis === 'row'
      ? { ...sourceRange, startRow: source[0]![0], endRow: source[0]![1] }
      : { ...sourceRange, startColumn: source[0]![0], endColumn: source[0]![1] }
    const mappedTarget = move.axis === 'row'
      ? { ...targetRange, startRow: target[0]![0], endRow: target[0]![1] }
      : { ...targetRange, startColumn: target[0]![0], endColumn: target[0]![1] }
    return { ops: [{ id: op.id, params: { ...p, [sourceKey]: mappedSource, [targetKey]: mappedTarget } }], lossy: false }
  }
  const p = op.params as Record<string, unknown>
  if (isRange(p.range)) {
    const mapped = mapRangeThroughMove(move, p.range)
    if (mapped.length !== 1) return { ops: [op], lossy: true }
    return { ops: [{ id: op.id, params: { ...p, range: mapped[0] } }], lossy: false }
  }
  if (Array.isArray(p.ranges) && p.ranges.every(isRange)) {
    const mapped = (p.ranges as IRangeLike[]).flatMap((range) => mapRangeThroughMove(move, range))
    return { ops: mapped.length ? [{ id: op.id, params: { ...p, ranges: mapped } }] : [], lossy: false }
  }
  return { ops: [op], lossy: false }
}

function remakeStruct(op: OpRecord, start: number, end: number): OpRecord {
  const p = op.params as Record<string, unknown> & { range: IRangeLike }
  const struct = parseStruct(op)!
  const range =
    struct.axis === 'row' ? { ...p.range, startRow: start, endRow: end } : { ...p.range, startColumn: start, endColumn: end }
  return { id: op.id, params: { ...p, range } }
}

export interface TransformResult {
  ops: OpRecord[]
  /** True when an untransformable op (a move, an unknown pair) passed through as-is. */
  lossy: boolean
}

/**
 * Transform `ops` through everything `against` does, in order. `against`
 * ops that aren't structural contribute nothing; moves and other
 * unmodelled structure flag the result lossy. A removed sheet consumes
 * every op that targeted it.
 */
export function transformOps(ops: OpRecord[], against: OpRecord[]): TransformResult {
  let cur = ops
  let lossy = false
  for (const a of against) {
    if (CONTENT_RANGE_OPS.has(a.id)) {
      // Content mutations write cells but do not change geometry. Cell-level
      // last-writer ownership is handled by reconcileRemote/withoutCells.
      lossy ||= cellsOf(a).size === 0
      continue
    }
    if (a.id === REMOVE_SHEET) {
      const gone = subUnitOf(a)
      cur = cur.filter((op) => subUnitOf(op) !== gone)
      continue
    }
    const move = parseMove(a)
    if (move) {
      const next: OpRecord[] = []
      for (const op of cur) {
        if (CONTENT_RANGE_OPS.has(op.id)) {
          next.push(op)
          lossy = true
          continue
        }
        const transformed = transformOpThroughMove(op, move)
        next.push(...transformed.ops)
        lossy ||= transformed.lossy
      }
      cur = next
      continue
    }
    const edit = parseStruct(a)
    if (!edit) continue
    const next: OpRecord[] = []
    for (const op of cur) {
      if (CONTENT_RANGE_OPS.has(op.id)) {
        next.push(op)
        lossy = true
        continue
      }
      const pendingMove = parseMove(op)
      if (pendingMove) {
        const transformed = transformMoveThroughStruct(op, pendingMove, edit)
        if (transformed.op) next.push(transformed.op)
        lossy ||= transformed.lossy
      } else {
        const transformed = transformOp(op, edit)
        if (transformed) next.push(transformed)
      }
    }
    cur = next
  }
  return { ops: cur, lossy }
}

/** Shift a peer's (or our own) selection through one structural edit. */
export function transformSelection(sel: SheetSelection, edit: StructEdit, editSheetId: string): SheetSelection {
  if (sel.sheet !== editSheetId) return sel
  const ranges: number[][] = []
  for (const [r0, c0, r1, c1] of sel.ranges) {
    const span = edit.axis === 'row' ? mapSpan(edit, r0, r1) : mapSpan(edit, c0, c1)
    if (!span) continue
    ranges.push(edit.axis === 'row' ? [span[0], c0, span[1], c1] : [r0, span[0], r1, span[1]])
  }
  const out: SheetSelection = {
    sheet: sel.sheet,
    ranges,
    ...(sel.mode ? { mode: sel.mode } : {}),
    ...(sel.draft !== undefined ? { draft: sel.draft } : {}),
  }
  if (sel.active) {
    const [r, c] = sel.active
    const m = edit.axis === 'row' ? mapIndex(edit, r) : mapIndex(edit, c)
    if (m !== null) out.active = edit.axis === 'row' ? [m, c] : [r, m]
  }
  return out
}

/** Permute a peer selection through a completed row/column move. */
export function transformSelectionThroughMove(sel: SheetSelection, move: MoveEdit): SheetSelection {
  if (sel.sheet !== move.subUnitId) return sel
  const ranges: number[][] = []
  for (const [r0, c0, r1, c1] of sel.ranges) {
    const spans = move.axis === 'row'
      ? mapSpanThroughMove(move, r0, r1)
      : mapSpanThroughMove(move, c0, c1)
    for (const span of spans) ranges.push(move.axis === 'row' ? [span[0], c0, span[1], c1] : [r0, span[0], r1, span[1]])
  }
  const out: SheetSelection = {
    sheet: sel.sheet,
    ranges,
    ...(sel.mode ? { mode: sel.mode } : {}),
    ...(sel.draft !== undefined ? { draft: sel.draft } : {}),
  }
  if (sel.active) {
    const [row, column] = sel.active
    out.active = move.axis === 'row'
      ? [mapIndexThroughMove(move, row), column]
      : [row, mapIndexThroughMove(move, column)]
  }
  return out
}

/** The structural edits (and lossiness) contained in one entry's ops. */
export function structEditsOf(ops: OpRecord[]): { edits: Array<{ edit: StructEdit }>; moves: MoveEdit[]; removedSheets: string[]; lossy: boolean } {
  const edits: Array<{ edit: StructEdit }> = []
  const moves: MoveEdit[] = []
  const removedSheets: string[] = []
  let lossy = false
  for (const op of ops) {
    if (CONTENT_RANGE_OPS.has(op.id)) lossy ||= cellsOf(op).size === 0
    else if (op.id === REMOVE_SHEET) {
      const s = subUnitOf(op)
      if (s) removedSheets.push(s)
    } else {
      const e = parseStruct(op)
      if (e) edits.push({ edit: e })
      else {
        const move = parseMove(op)
        if (move) moves.push(move)
        else if (op.id === 'sheet.mutation.move-rows' || op.id === 'sheet.mutation.move-columns') lossy = true
      }
    }
  }
  return { edits, moves, removedSheets, lossy }
}
