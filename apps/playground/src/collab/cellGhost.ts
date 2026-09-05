import { colName, type PeerInfo, type SheetSelection } from '../../../../packages/collab/src/index.js'

/** A rectangle in coordinates local to the editor stage. */
export interface CellGhostRect {
  left: number
  top: number
  width: number
  height: number
}

export interface CellGhostTarget {
  sheet: string
  row: number
  column: number
  a1: string
  draft: string
}

export interface CellGhostViewModel extends CellGhostTarget {
  name: string
  color: string
  rect: CellGhostRect
  role: 'status'
  ariaLive: 'polite'
  ariaAtomic: true
  ariaLabel: string
  animate: boolean
}

function finiteRect(rect: CellGhostRect): boolean {
  return Number.isFinite(rect.left)
    && Number.isFinite(rect.top)
    && Number.isFinite(rect.width)
    && Number.isFinite(rect.height)
    && rect.width > 0
    && rect.height > 0
}

/**
 * Resolve the one cell that owns an ephemeral draft. An empty draft, a
 * committed/cancelled edit, or a selection on another sheet has no ghost.
 */
export function cellGhostTarget(
  selection: SheetSelection | null | undefined,
  activeSheet: string,
): CellGhostTarget | null {
  if (!selection
    || selection.sheet !== activeSheet
    || selection.mode !== 'editing'
    || typeof selection.draft !== 'string'
    || selection.draft.length === 0
    || selection.ranges.length === 0) return null

  const [row, column] = selection.active ?? [selection.ranges[0][0], selection.ranges[0][1]]
  if (!Number.isInteger(row) || row < 0 || !Number.isInteger(column) || column < 0) return null
  return {
    sheet: selection.sheet,
    row,
    column,
    a1: `${colName(column)}${row + 1}`,
    draft: selection.draft,
  }
}

/** Return the visible part of a cell, or null when scrolling clips it away. */
export function clipCellGhostRect(cell: CellGhostRect, viewport: CellGhostRect): CellGhostRect | null {
  if (!finiteRect(cell) || !finiteRect(viewport)) return null
  const left = Math.max(cell.left, viewport.left)
  const top = Math.max(cell.top, viewport.top)
  const right = Math.min(cell.left + cell.width, viewport.left + viewport.width)
  const bottom = Math.min(cell.top + cell.height, viewport.top + viewport.height)
  if (right <= left || bottom <= top) return null
  return { left, top, width: right - left, height: bottom - top }
}

/**
 * Build presentation and accessibility state from independently observed
 * selection and geometry. Re-running this after scroll/resize is intentional:
 * no stale coordinates are kept here.
 */
export function cellGhostViewModel(
  peer: Pick<PeerInfo, 'name' | 'color'>,
  selection: SheetSelection | null | undefined,
  activeSheet: string,
  cell: CellGhostRect | null,
  viewport: CellGhostRect,
  prefersReducedMotion = false,
): CellGhostViewModel | null {
  const target = cellGhostTarget(selection, activeSheet)
  const rect = cell && clipCellGhostRect(cell, viewport)
  if (!target || !rect) return null
  return {
    ...target,
    name: peer.name,
    color: peer.color,
    rect,
    role: 'status',
    ariaLive: 'polite',
    ariaAtomic: true,
    ariaLabel: `${peer.name} is editing ${target.a1}: ${target.draft}`,
    animate: !prefersReducedMotion,
  }
}
