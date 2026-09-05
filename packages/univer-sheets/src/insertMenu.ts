import {
  CommandType,
  ICommandService,
  type IDisposable,
  type Univer,
} from '@univerjs/core'
import {
  IMenuManagerService,
  MenuItemType,
  RibbonInsertGroup,
  RibbonPosition,
} from '@univerjs/ui'

export const INJOFFICE_INSERT_COMMANDS = {
  chart: 'injoffice.command.insert-chart',
  pivotTable: 'injoffice.command.insert-pivot-table',
  shape: 'injoffice.command.insert-shape',
} as const

export interface InjOfficeInsertActions {
  chart?: () => boolean | void | Promise<boolean | void>
  pivotTable?: () => boolean | void | Promise<boolean | void>
  shape?: () => boolean | void | Promise<boolean | void>
}

export type InjOfficeInsertFeatureConfig = Partial<Record<keyof InjOfficeInsertActions, boolean>>

/**
 * Put independently implemented InjOffice objects into Univer's Insert ribbon.
 * A missing action or a `false` feature flag omits the command and menu item.
 */
export function registerInjOfficeInsertMenu(
  univer: Univer,
  actions: InjOfficeInsertActions,
  features: InjOfficeInsertFeatureConfig = {},
): IDisposable {
  const injector = univer.__getInjector()
  const commandService = injector.get(ICommandService)
  const menuManager = injector.get(IMenuManagerService)
  const registrations: IDisposable[] = []

  const entries = [
    ['chart', 'Chart', 10],
    ['pivotTable', 'Pivot table', 20],
    ['shape', 'Shape', 30],
  ] as const

  for (const [key, title, order] of entries) {
    const action = actions[key]
    if (!action || features[key] === false) continue
    const commandId = INJOFFICE_INSERT_COMMANDS[key]
    registrations.push(commandService.registerCommand({
      id: commandId,
      type: CommandType.COMMAND,
      handler: async () => (await action()) !== false,
    }))
    menuManager.mergeMenu({
      [RibbonPosition.INSERT]: {
        [RibbonInsertGroup.MEDIA]: {
          [commandId]: {
            order,
            menuItemFactory: () => ({
              id: commandId,
              title,
              tooltip: `Insert ${title.toLowerCase()}`,
              type: MenuItemType.BUTTON,
            }),
          },
        },
      },
    })
  }

  return {
    dispose() {
      for (const registration of registrations.reverse()) registration.dispose()
    },
  }
}
