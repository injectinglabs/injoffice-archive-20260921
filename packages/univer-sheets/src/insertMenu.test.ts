import type { Univer } from '@univerjs/core'
import { describe, expect, it, vi } from 'vitest'
import { INJOFFICE_INSERT_COMMANDS, registerInjOfficeInsertMenu } from './insertMenu'

function harness() {
  const commands: Array<{ id: string; handler: () => unknown }> = []
  const menus: unknown[] = []
  const dispose = vi.fn()
  const services = [
    { registerCommand: (command: { id: string; handler: () => unknown }) => { commands.push(command); return { dispose } } },
    { mergeMenu: (menu: unknown) => menus.push(menu) },
  ]
  let index = 0
  const univer = { __getInjector: () => ({ get: () => services[index++] }) } as unknown as Univer
  return { commands, dispose, menus, univer }
}

describe('InjOffice Insert ribbon bridge', () => {
  it('registers only supplied and enabled actions', async () => {
    const subject = harness()
    const chart = vi.fn()
    const shape = vi.fn()
    registerInjOfficeInsertMenu(subject.univer, { chart, shape }, { shape: false })

    expect(subject.commands.map((command) => command.id)).toEqual([INJOFFICE_INSERT_COMMANDS.chart])
    expect(subject.menus).toHaveLength(1)
    await subject.commands[0].handler()
    expect(chart).toHaveBeenCalledOnce()
    expect(shape).not.toHaveBeenCalled()
  })

  it('unregisters all commands when disposed', () => {
    const subject = harness()
    const registration = registerInjOfficeInsertMenu(subject.univer, {
      chart: () => true,
      pivotTable: () => true,
      shape: () => true,
    })
    registration.dispose()
    expect(subject.dispose).toHaveBeenCalledTimes(3)
  })
})
