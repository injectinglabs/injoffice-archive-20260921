import type { SheetSelection } from './types'
import { cellsOf, MOVE_RANGE, REORDER_RANGE, SET_RANGE_VALUES, withoutCells, type OpRecord } from './sync'
import {
  parseMove,
  parseStruct,
  structEditsOf,
  transformOps,
  transformSelection,
  transformSelectionThroughMove,
} from './transform'

type Direction = 'undo' | 'redo'

export interface CollaborativeUndoEntry {
  /** Stable host identity. A deterministic package-local id is assigned when omitted. */
  id?: string
  /** Mutations that reverse the local action, in execution order. */
  undoOps: OpRecord[]
  /** Mutations that repeat the local action, in execution order. */
  redoOps: OpRecord[]
  selectionBefore?: SheetSelection | null
  selectionAfter?: SheetSelection | null
}

export type UndoBlockCode =
  | 'unsupported-local-operation'
  | 'unsupported-remote-operation'
  | 'lossy-structural-transform'
  | 'remote-write-after-local-structure'

export interface UndoBlockReason {
  code: UndoBlockCode
  opId: string
  message: string
}

export type UndoPlan =
  | { status: 'empty'; direction: Direction }
  | { status: 'blocked'; direction: Direction; entryId: string; reasons: readonly UndoBlockReason[] }
  | {
      status: 'ready' | 'neutralized'
      direction: Direction
      entryId: string
      /** Invalidates a plan when record(), clear(), or rebaseRemote() intervenes. */
      revision: number
      ops: readonly OpRecord[]
      selection: SheetSelection | null
    }

export interface UndoRebaseReport {
  rebased: number
  neutralized: number
  blocked: number
}

export interface UndoRebaseOptions {
  /** Entries whose local mutation is still outbound and will be ordered after
   * this remote batch. Their geometry moves, but earlier remote cell writes
   * must not steal later local ownership. */
  remoteBeforeEntryIds?: readonly string[]
}

interface StoredEntry {
  id: string
  undoOps: OpRecord[]
  redoOps: OpRecord[]
  selectionBefore: SheetSelection | null
  selectionAfter: SheetSelection | null
  reasons: UndoBlockReason[]
}

const CONTENT_RANGE_OPS = new Set([MOVE_RANGE, REORDER_RANGE])
const STRUCTURAL_IDS = new Set([
  'sheet.mutation.insert-row',
  'sheet.mutation.insert-col',
  'sheet.mutation.remove-rows',
  'sheet.mutation.remove-col',
  'sheet.mutation.move-rows',
  'sheet.mutation.move-columns',
  'sheet.mutation.remove-sheet',
])
const OBJECT_ID = /(?:drawing|image|shape|chart|floating-object|rich-text)/i

/**
 * A host-neutral collaborative undo stack.
 *
 * The host records already-computed forward/inverse mutation batches. Remote
 * batches then rebase both stacks before the next plan is requested. Planning
 * and committing are deliberately separate: a host executes `plan.ops` as a
 * normal local mutation batch (so it is synchronized), then commits only after
 * execution succeeds. Any intervening remote rebase invalidates the plan.
 */
export class CollaborativeUndoManager {
  private undoStack: StoredEntry[] = []
  private redoStack: StoredEntry[] = []
  private revision = 0
  private nextId = 1

  get undoDepth(): number {
    return this.undoStack.length
  }

  get redoDepth(): number {
    return this.redoStack.length
  }

  record(input: CollaborativeUndoEntry): string {
    const id = input.id ?? `collab-undo-${this.nextId++}`
    if (!id) throw new Error('Collaborative undo entry id must not be empty')
    if (this.undoStack.some((entry) => entry.id === id) || this.redoStack.some((entry) => entry.id === id)) {
      throw new Error(`Collaborative undo entry id already exists: ${id}`)
    }
    const undoOps = cloneOps(input.undoOps)
    const redoOps = cloneOps(input.redoOps)
    const stored: StoredEntry = {
      id,
      undoOps,
      redoOps,
      selectionBefore: cloneSelection(input.selectionBefore),
      selectionAfter: cloneSelection(input.selectionAfter),
      reasons: [],
    }
    addReasons(stored, unsupportedReasons([...undoOps, ...redoOps], 'local'))
    this.undoStack.push(stored)
    this.redoStack = []
    this.revision++
    return id
  }

  clear(): void {
    this.undoStack = []
    this.redoStack = []
    this.revision++
  }

  planUndo(): UndoPlan {
    return this.plan('undo')
  }

  planRedo(): UndoPlan {
    return this.plan('redo')
  }

  /**
   * Commit a successfully executed plan. Returns false for a stale plan or
   * when the planned entry is no longer at the top of the relevant stack.
   */
  commit(plan: UndoPlan): boolean {
    if (plan.status !== 'ready' && plan.status !== 'neutralized') return false
    if (plan.revision !== this.revision) return false
    const source = plan.direction === 'undo' ? this.undoStack : this.redoStack
    const target = plan.direction === 'undo' ? this.redoStack : this.undoStack
    if (source.at(-1)?.id !== plan.entryId) return false
    target.push(source.pop()!)
    this.revision++
    return true
  }

  /** Re-express every local history entry on top of an ordered remote batch. */
  rebaseRemote(remoteOps: readonly OpRecord[], options: UndoRebaseOptions = {}): UndoRebaseReport {
    if (remoteOps.length === 0) return this.report()
    const remote = cloneOps(remoteOps)
    const remoteBefore = new Set(options.remoteBeforeEntryIds ?? [])
    for (const entry of [...this.undoStack, ...this.redoStack]) {
      this.rebaseEntry(entry, remote, remoteBefore.has(entry.id))
    }
    this.revision++
    return this.report()
  }

  private plan(direction: Direction): UndoPlan {
    const entry = (direction === 'undo' ? this.undoStack : this.redoStack).at(-1)
    if (!entry) return { status: 'empty', direction }
    if (entry.reasons.length > 0) return { status: 'blocked', direction, entryId: entry.id, reasons: cloneReasons(entry.reasons) }
    const ops = cloneOps(direction === 'undo' ? entry.undoOps : entry.redoOps)
    return {
      status: ops.length === 0 ? 'neutralized' : 'ready',
      direction,
      entryId: entry.id,
      revision: this.revision,
      ops,
      selection: cloneSelection(direction === 'undo' ? entry.selectionBefore : entry.selectionAfter),
    }
  }

  private rebaseEntry(entry: StoredEntry, remoteOps: OpRecord[], remoteIsEarlier: boolean): void {
    for (const remote of remoteOps) {
      const unsupported = unsupportedReasons([remote], 'remote')
      if (unsupported.length > 0 && intersectsEntry(entry, remote)) {
        addReasons(entry, unsupported)
        continue
      }

      if (cellsOf(remote).size > 0) {
        if (remoteIsEarlier) continue
        if (hasStructuralOps(entry) && intersectsEntry(entry, remote)) {
          addReasons(entry, [{
            code: 'remote-write-after-local-structure',
            opId: remote.id,
            message: 'A later remote cell write may be owned by rows or columns created or removed by this local action.',
          }])
          continue
        }
        entry.undoOps = dropOverwrittenCells(entry.undoOps, remote)
        entry.redoOps = dropOverwrittenCells(entry.redoOps, remote)
        continue
      }

      const undo = transformOps(entry.undoOps, [remote])
      const redo = transformOps(entry.redoOps, [remote])
      entry.undoOps = undo.ops
      entry.redoOps = redo.ops
      if (undo.lossy || redo.lossy) {
        addReasons(entry, [{
          code: 'lossy-structural-transform',
          opId: remote.id,
          message: 'The remote structural operation would split or ambiguously remap this local undo entry.',
        }])
      }
      entry.selectionBefore = mapSelection(entry.selectionBefore, remote)
      entry.selectionAfter = mapSelection(entry.selectionAfter, remote)
    }
  }

  private report(): UndoRebaseReport {
    const entries = [...this.undoStack, ...this.redoStack]
    return {
      rebased: entries.filter((entry) => entry.reasons.length === 0 && (entry.undoOps.length > 0 || entry.redoOps.length > 0)).length,
      neutralized: entries.filter((entry) => entry.reasons.length === 0 && entry.undoOps.length === 0 && entry.redoOps.length === 0).length,
      blocked: entries.filter((entry) => entry.reasons.length > 0).length,
    }
  }
}

function cloneOps(ops: readonly OpRecord[]): OpRecord[] {
  return ops.map((op) => ({ id: op.id, params: JSON.parse(JSON.stringify(op.params ?? {})) as Record<string, unknown> }))
}

function cloneSelection(selection: SheetSelection | null | undefined): SheetSelection | null {
  if (!selection) return null
  return {
    ...selection,
    ranges: selection.ranges.map((range) => [...range]),
    ...(selection.active ? { active: [...selection.active] as [number, number] } : {}),
  }
}

function cloneReasons(reasons: UndoBlockReason[]): UndoBlockReason[] {
  return reasons.map((reason) => ({ ...reason }))
}

function addReasons(entry: StoredEntry, reasons: UndoBlockReason[]): void {
  for (const reason of reasons) {
    if (!entry.reasons.some((existing) => existing.code === reason.code && existing.opId === reason.opId)) entry.reasons.push(reason)
  }
}

function unsupportedReasons(ops: readonly OpRecord[], source: 'local' | 'remote'): UndoBlockReason[] {
  const code = source === 'local' ? 'unsupported-local-operation' : 'unsupported-remote-operation'
  const reasons: UndoBlockReason[] = []
  for (const op of ops) {
    let message = ''
    const struct = parseStruct(op)
    if (CONTENT_RANGE_OPS.has(op.id) && cellsOf(op).size === 0) message = 'Content-range undo requires valid absolute move matrices or a bounded row-reorder range and order map.'
    else if (OBJECT_ID.test(op.id)) message = 'Rich and floating object mutations require stable object identity and object-specific inverse transforms.'
    else if (
      STRUCTURAL_IDS.has(op.id)
      && op.id !== 'sheet.mutation.remove-sheet'
      && ((!struct && !parseMove(op)) || (struct && (!Number.isSafeInteger(struct.start) || struct.start < 0 || !Number.isSafeInteger(struct.count) || struct.count < 1)))
    ) {
      message = 'The structural mutation is malformed or outside the supported contiguous row/column vocabulary.'
    } else if (op.id === SET_RANGE_VALUES && !validCellMatrix(op.params.cellValue)) {
      message = 'Cell undo requires a sparse matrix with non-negative integer row and column keys.'
    }
    if (message) reasons.push({ code, opId: op.id, message })
  }
  return reasons
}

function validCellMatrix(value: unknown): boolean {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  for (const [rowKey, row] of Object.entries(value as Record<string, unknown>)) {
    if (!validIndex(rowKey) || !row || typeof row !== 'object' || Array.isArray(row)) return false
    for (const columnKey of Object.keys(row as Record<string, unknown>)) if (!validIndex(columnKey)) return false
  }
  return true
}

function validIndex(value: string): boolean {
  const number = Number(value)
  return Number.isSafeInteger(number) && number >= 0 && String(number) === value
}

function subUnitOf(op: OpRecord): string | null {
  return typeof op.params.subUnitId === 'string' ? op.params.subUnitId : null
}

function entrySheets(entry: StoredEntry): Set<string> {
  const sheets = new Set<string>()
  for (const op of [...entry.undoOps, ...entry.redoOps]) {
    const sheet = subUnitOf(op)
    if (sheet) sheets.add(sheet)
  }
  if (entry.selectionBefore) sheets.add(entry.selectionBefore.sheet)
  if (entry.selectionAfter) sheets.add(entry.selectionAfter.sheet)
  return sheets
}

function intersectsEntry(entry: StoredEntry, op: OpRecord): boolean {
  const sheet = subUnitOf(op)
  return sheet === null || entrySheets(entry).has(sheet)
}

function hasStructuralOps(entry: StoredEntry): boolean {
  return [...entry.undoOps, ...entry.redoOps].some((op) => STRUCTURAL_IDS.has(op.id))
}

function dropOverwrittenCells(localOps: OpRecord[], remote: OpRecord): OpRecord[] {
  const overwritten = cellsOf(remote)
  return localOps.flatMap((op) => withoutCells(op, overwritten))
}

function mapSelection(selection: SheetSelection | null, remote: OpRecord): SheetSelection | null {
  if (!selection) return null
  const { edits, moves, removedSheets } = structEditsOf([remote])
  if (removedSheets.includes(selection.sheet)) return null
  let mapped = selection
  for (const { edit } of edits) mapped = transformSelection(mapped, edit, edit.subUnitId)
  for (const move of moves) mapped = transformSelectionThroughMove(mapped, move)
  return mapped
}
