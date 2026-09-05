import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it } from 'vitest'
import { OutlineCommandController } from './commands'
import { OutlineManager } from './manager'
import { OutlineGutter, createOutlineGutterCommandBindings } from './react'
import type { OutlineGroup } from './types'

const visibility = { hide() {}, show() {} }

function group(id: string, start: number, end: number, collapsed = false): OutlineGroup {
  return { id, sheetId: 's', axis: 'row', start, end, collapsed }
}

describe('OutlineGutter', () => {
  it('routes expand, collapse, toggle, and level controls through the command facade', () => {
    const records: Array<{ label: string }> = []
    const controller = new OutlineCommandController(new OutlineManager(visibility), {
      push(record) { records.push(record) },
    })
    controller.add(group('outer', 0, 9))
    controller.add(group('inner', 2, 5))
    records.length = 0
    const commands = createOutlineGutterCommandBindings(controller)
    expect(commands.collapse('outer').ok).toBe(true)
    expect(commands.expand('outer').ok).toBe(true)
    expect(commands.toggle('inner').ok).toBe(true)
    expect(commands.showLevel('s', 'row', 1)).toBe(1)
    expect(records.map(({ label }) => label)).toEqual([
      'Collapse outline',
      'Expand outline',
      'Collapse outline',
      'Collapse outline',
    ])
  })

  it('renders an accessible margin with level and expand/collapse controls', () => {
    const controller = new OutlineCommandController(new OutlineManager(visibility))
    controller.add(group('outer', 0, 9))
    controller.add(group('inner', 2, 5, true))
    const markup = renderToStaticMarkup(<OutlineGutter controller={controller} sheetId="s" axis="row" />)
    expect(markup).toContain('aria-label="row outline margin"')
    expect(markup).toContain('aria-label="row outline levels"')
    expect(markup).toContain('aria-label="Show outline level 1"')
    expect(markup).toContain('aria-label="Show outline level 2"')
    expect(markup).toContain('aria-label="Collapse rows 1–10"')
    expect(markup).toContain('aria-label="Expand rows 3–6"')
    expect(markup).toContain('aria-expanded="true"')
    expect(markup).toContain('aria-expanded="false"')
  })
})
