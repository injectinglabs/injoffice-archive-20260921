import { describe, expect, it } from 'vitest'
import { SparklineCommandController } from './commands'
import { SparklineManager } from './manager'

const input = (id: string, column: number) => ({
  id,
  type: 'line' as const,
  source: { sheetId: 's', startRow: 0, startColumn: 0, endRow: 0, endColumn: 2 },
  target: { sheetId: 's', row: 0, column },
})

describe('SparklineCommandController', () => {
  it('provides live handles and one inverse for grouped lifecycle changes', () => {
    const records: any[] = []
    const subject = new SparklineCommandController(new SparklineManager(), { push: (record) => records.push(record) })
    const first = subject.create(input('one', 3))
    subject.create(input('two', 4))
    subject.group(['one', 'two'], 'group')
    expect(first.value?.groupId).toBe('group')
    expect(records).toHaveLength(3)
    expect(records[2]).toMatchObject({
      label: 'Group sparklines',
      before: { groups: [] },
      after: { groups: [{ id: 'group', memberIds: ['one', 'two'] }] },
    })
  })

  it('fails closed when restoring an invalid snapshot', () => {
    const manager = new SparklineManager()
    const subject = new SparklineCommandController(manager)
    subject.create(input('one', 3))
    const before = manager.serialize()
    expect(subject.restore({ version: 1, sparklines: [], groups: [{ id: 'orphan', memberIds: ['x', 'y'] }] })).toBe(false)
    expect(manager.serialize()).toEqual(before)
  })
})
