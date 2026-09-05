import { describe, expect, it } from 'vitest'
import { OutlineCommandController } from './commands'
import { OutlineManager } from './manager'
import { applyOutlineLevel, collapseToOutlineLevel, layoutOutlineGutter, outlineGroupDepth } from './gutter'
import type { OutlineGroup } from './types'

const visibility = { hide() {}, show() {} }

function group(id: string, start: number, end: number, collapsed = false): OutlineGroup {
  return { id, sheetId: 's', axis: 'row', start, end, collapsed }
}

describe('outline gutter layout', () => {
  it('orders nested groups and reports 1-based depth', () => {
    const groups = [group('outer', 0, 9), group('inner', 2, 5), group('other', 12, 14)]
    expect(outlineGroupDepth(groups, groups[1]!)).toBe(2)
    const model = layoutOutlineGutter(groups, 's', 'row')
    expect(model.maxLevel).toBe(2)
    expect(model.controls.map(({ id, depth }) => [id, depth])).toEqual([
      ['outer', 1],
      ['inner', 2],
      ['other', 1],
    ])
    expect(layoutOutlineGutter(groups, 's', 'column').controls).toEqual([])
  })

  it('maps Excel-style level buttons onto collapse state', () => {
    const model = layoutOutlineGutter([group('outer', 0, 9), group('inner', 2, 5)], 's', 'row')
    expect(collapseToOutlineLevel(model, 1)).toEqual([
      { id: 'outer', collapsed: true },
      { id: 'inner', collapsed: true },
    ])
    expect(collapseToOutlineLevel(model, 2)).toEqual([
      { id: 'outer', collapsed: false },
      { id: 'inner', collapsed: true },
    ])
    expect(collapseToOutlineLevel(model, 3)).toEqual([
      { id: 'outer', collapsed: false },
      { id: 'inner', collapsed: false },
    ])
    expect(() => collapseToOutlineLevel(model, 0)).toThrow('outline level must be a positive integer')
  })

  it('applies a level through snapshot undo without rewriting unchanged groups', () => {
    const records: Array<{ label: string }> = []
    const controller = new OutlineCommandController(new OutlineManager(visibility), {
      push(record) { records.push(record) },
    })
    controller.add(group('outer', 0, 9))
    controller.add(group('inner', 2, 5, true))
    records.length = 0
    expect(applyOutlineLevel(controller, 's', 'row', 2)).toBe(0)
    expect(applyOutlineLevel(controller, 's', 'row', 1)).toBe(1)
    expect(controller.manager.get('outer')?.collapsed).toBe(true)
    expect(controller.manager.get('inner')?.collapsed).toBe(true)
    expect(records.map(({ label }) => label)).toEqual(['Collapse outline'])
  })
})
