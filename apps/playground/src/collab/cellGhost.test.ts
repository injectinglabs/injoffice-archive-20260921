import { describe, expect, it } from 'vitest'
import {
  cellGhostTarget,
  cellGhostViewModel,
  clipCellGhostRect,
  type CellGhostRect,
} from './cellGhost'

const viewport: CellGhostRect = { left: 42, top: 28, width: 560, height: 320 }
const cell: CellGhostRect = { left: 120, top: 84, width: 105, height: 24 }
const peer = { name: 'Mira', color: '#2855d9' }

describe('cellGhostTarget', () => {
  it('is visible only for a nonempty draft in editing mode on the active sheet', () => {
    const editing = { sheet: 's1', ranges: [[4, 3, 4, 3]], active: [4, 3] as [number, number], mode: 'editing' as const, draft: 'Ready' }
    expect(cellGhostTarget(editing, 's1')).toMatchObject({ row: 4, column: 3, a1: 'D5', draft: 'Ready' })
    expect(cellGhostTarget({ ...editing, mode: 'selecting' }, 's1')).toBeNull()
    expect(cellGhostTarget({ ...editing, draft: '' }, 's1')).toBeNull()
    expect(cellGhostTarget({ ...editing, draft: undefined }, 's1')).toBeNull()
    expect(cellGhostTarget(editing, 'other-sheet')).toBeNull()
  })

  it('targets active row/column and falls back to the first range start', () => {
    expect(cellGhostTarget({
      sheet: 's1',
      ranges: [[2, 1, 8, 4]],
      active: [12, 26],
      mode: 'editing',
      draft: 'Active cell wins',
    }, 's1')).toMatchObject({ row: 12, column: 26, a1: 'AA13' })

    expect(cellGhostTarget({
      sheet: 's1',
      ranges: [[7, 2, 9, 5]],
      mode: 'editing',
      draft: 'Range fallback',
    }, 's1')).toMatchObject({ row: 7, column: 2, a1: 'C8' })
  })

  it('cleans up immediately after commit, cancel, or peer removal', () => {
    const live = { sheet: 's1', ranges: [[1, 1, 1, 1]], mode: 'editing' as const, draft: 'typing' }
    expect(cellGhostTarget(live, 's1')).not.toBeNull()
    // Univer emits selecting after either accepting or cancelling its cell editor.
    expect(cellGhostTarget({ ...live, mode: 'selecting', draft: undefined }, 's1')).toBeNull()
    expect(cellGhostTarget(null, 's1')).toBeNull()
  })
})

describe('cell ghost geometry', () => {
  it('clips a partially visible cell to the editor viewport', () => {
    expect(clipCellGhostRect({ left: 20, top: 20, width: 40, height: 30 }, viewport)).toEqual({
      left: 42,
      top: 28,
      width: 18,
      height: 22,
    })
    expect(clipCellGhostRect({ left: 700, top: 84, width: 80, height: 24 }, viewport)).toBeNull()
  })

  it('recomputes its position after scrolling instead of retaining stale coordinates', () => {
    const selection = { sheet: 's1', ranges: [[4, 3, 4, 3]], mode: 'editing' as const, draft: 'Live' }
    const before = cellGhostViewModel(peer, selection, 's1', cell, viewport)
    const after = cellGhostViewModel(peer, selection, 's1', { ...cell, left: 68, top: 46 }, viewport)
    expect(before?.rect).toEqual(cell)
    expect(after?.rect).toEqual({ left: 68, top: 46, width: 105, height: 24 })
    expect(after?.row).toBe(before?.row)
    expect(after?.column).toBe(before?.column)
  })
})

describe('cell ghost accessibility', () => {
  const selection = { sheet: 's1', ranges: [[4, 3, 4, 3]], mode: 'editing' as const, draft: 'Ready' }

  it('announces collaborator, target, and draft without stealing focus', () => {
    expect(cellGhostViewModel(peer, selection, 's1', cell, viewport)).toMatchObject({
      role: 'status',
      ariaLive: 'polite',
      ariaAtomic: true,
      ariaLabel: 'Mira is editing D5: Ready',
    })
  })

  it('disables decorative movement when reduced motion is requested', () => {
    expect(cellGhostViewModel(peer, selection, 's1', cell, viewport, false)?.animate).toBe(true)
    expect(cellGhostViewModel(peer, selection, 's1', cell, viewport, true)?.animate).toBe(false)
  })
})
