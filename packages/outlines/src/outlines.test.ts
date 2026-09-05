import { describe, expect, it, vi } from 'vitest'
import { OutlineManager } from './manager'
import { OutlineStore } from './store'
import { transformOutline } from './transform'
import { validateOutlineGroup } from './validation'
import type { OutlineGroup, OutlineVisibilityAdapter } from './types'

const row = (id: string, start: number, end: number, collapsed = false): OutlineGroup => ({
  id,
  sheetId: 'sheet-1',
  axis: 'row',
  start,
  end,
  collapsed,
})

describe('outline validation and storage', () => {
  it('allows disjoint and nested groups while refusing crossing ranges', () => {
    const outer = row('outer', 1, 10)
    expect(validateOutlineGroup(row('inner', 2, 5), [outer]).ok).toBe(true)
    expect(validateOutlineGroup(row('other', 12, 15), [outer]).ok).toBe(true)
    const crossing = validateOutlineGroup(row('crossing', 8, 12), [outer])
    expect(crossing.ok).toBe(false)
    if (!crossing.ok) expect(crossing.issues.map((issue) => issue.code)).toContain('CROSSING')
  })

  it('hydrates atomically and reports hierarchy depth', () => {
    const store = new OutlineStore()
    expect(store.hydrate([row('outer', 1, 10), row('inner', 2, 5)])).toMatchObject({ ok: true })
    expect(store.depth('outer')).toBe(1)
    expect(store.depth('inner')).toBe(2)
    expect(store.hydrate([row('a', 1, 5), row('b', 4, 8)])).toMatchObject({ ok: false })
    expect(store.list().map((group) => group.id)).toEqual(['outer', 'inner'])
  })
})

describe('outline visibility manager', () => {
  it('keeps an inner collapsed group hidden when its parent expands', () => {
    const visibility: OutlineVisibilityAdapter = { hide: vi.fn(), show: vi.fn() }
    const manager = new OutlineManager(visibility)
    manager.add(row('outer', 1, 10, true))
    manager.add(row('inner', 3, 5, true))
    manager.setCollapsed('outer', false)
    expect(visibility.show).toHaveBeenLastCalledWith('sheet-1', 'row', 1, 10)
    expect(visibility.hide).toHaveBeenLastCalledWith('sheet-1', 'row', 3, 3)
  })

  it('emits changes only for successful mutations', () => {
    const manager = new OutlineManager({ hide: vi.fn(), show: vi.fn() })
    const listener = vi.fn()
    manager.onChange(listener)
    expect(manager.add(row('group', 2, 4)).ok).toBe(true)
    expect(manager.remove('missing')).toBe(false)
    expect(manager.toggle('group').ok).toBe(true)
    expect(listener).toHaveBeenCalledTimes(2)
  })
})

describe('outline structural transforms', () => {
  it('shifts, expands, shrinks, and removes outline ranges', () => {
    expect(transformOutline(row('g', 4, 8), { kind: 'insert', axis: 'row', sheetId: 'sheet-1', start: 2, count: 2 }))
      .toMatchObject({ start: 6, end: 10 })
    expect(transformOutline(row('g', 4, 8), { kind: 'insert', axis: 'row', sheetId: 'sheet-1', start: 6, count: 2 }))
      .toMatchObject({ start: 4, end: 10 })
    expect(transformOutline(row('g', 4, 8), { kind: 'remove', axis: 'row', sheetId: 'sheet-1', start: 6, count: 2 }))
      .toMatchObject({ start: 4, end: 6 })
    expect(transformOutline(row('g', 4, 8), { kind: 'remove', axis: 'row', sheetId: 'sheet-1', start: 2, count: 10 }))
      .toBeNull()
  })
})
