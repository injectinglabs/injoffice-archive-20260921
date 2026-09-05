import { describe, expect, it, vi } from 'vitest'
import { OUTLINE_COLLABORATION_APPLY, OutlineCommandController, type OutlineUndoRecord } from './commands'
import { OutlineManager } from './manager'

const visibility = () => ({ hide: vi.fn(), show: vi.fn() })

describe('OutlineCommandController', () => {
  it('exposes live handles and complete inverse snapshots', () => {
    const records: unknown[] = []
    const subject = new OutlineCommandController(new OutlineManager(visibility()), { push: (record) => records.push(record) })
    expect(subject.add({ id: 'rows', sheetId: 's', axis: 'row', start: 2, end: 5, collapsed: false }).ok).toBe(true)
    const handle = subject.getById('rows')!
    expect(handle.value?.collapsed).toBe(false)
    subject.setCollapsed('rows', true)
    expect(handle.value?.collapsed).toBe(true)
    expect(records).toMatchObject([
      { label: 'Add outline', before: [], after: [{ id: 'rows' }] },
      { label: 'Collapse outline', before: [{ collapsed: false }], after: [{ collapsed: true }] },
    ])
  })

  it('restores snapshots atomically and rejects malformed payloads', () => {
    const manager = new OutlineManager(visibility())
    const subject = new OutlineCommandController(manager)
    subject.add({ id: 'rows', sheetId: 's', axis: 'row', start: 2, end: 5, collapsed: false })
    const before = manager.serialize()
    expect(subject.restore([{ id: 'bad', sheetId: 's', axis: 'row', start: 5, end: 2, collapsed: false }])).toBe(false)
    expect(manager.serialize()).toEqual(before)
    expect(subject.restore([])).toBe(true)
    expect(manager.serialize()).toEqual([])
  })

  it('applies authoritative snapshots with explicit undo policy in collaboration mode', () => {
    const records: OutlineUndoRecord[] = []
    const manager = new OutlineManager(visibility(), { mutationAuthority: 'collaboration' })
    const subject = new OutlineCommandController(manager, { push: (record) => records.push(record) })
    expect(subject.add({ id: 'blocked', sheetId: 's', axis: 'row', start: 1, end: 2, collapsed: false }).ok).toBe(false)
    expect(subject[OUTLINE_COLLABORATION_APPLY]([{ id: 'rows', sheetId: 's', axis: 'row', start: 1, end: 2, collapsed: true }], { label: 'Remote add', recordUndo: true })).toBe(true)
    expect(records).toHaveLength(1)
    expect(records[0]?.label).toBe('Remote add')
    expect(subject[OUTLINE_COLLABORATION_APPLY]([], { recordUndo: false })).toBe(true)
    expect(records).toHaveLength(1)
  })

  it('rolls the model back when authoritative visibility application fails', () => {
    const manager = new OutlineManager({ hide: () => { throw new Error('renderer failed') }, show: vi.fn() }, { mutationAuthority: 'collaboration' })
    const subject = new OutlineCommandController(manager)
    expect(subject[OUTLINE_COLLABORATION_APPLY]([{ id: 'rows', sheetId: 's', axis: 'row', start: 1, end: 2, collapsed: true }])).toBe(false)
    expect(subject.snapshot()).toEqual([])
  })
})
