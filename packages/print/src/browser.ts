import { CommandType, ICommandService, IUndoRedoService, type IDisposable, type Univer } from '@univerjs/core'
import { IMenuManagerService, MenuItemType, RibbonPosition, RibbonStartGroup } from '@univerjs/ui'
import type { Observable } from 'rxjs'
import { PrintConfigurationCommandController, type PrintConfigurationCommandTarget, type PrintConfigurationPersistence } from './commands'
import { PrintManager } from './manager'
import type { PrintConfigurationSnapshotV1, PrintLayoutConfig, PrintRenderConfig } from './types'

export * from './index'

export const INJOFFICE_PRINT_CONFIGURATION_COMMANDS = {
  restore: 'injoffice.mutation.print-configuration.restore',
  updateLayout: 'injoffice.command.print-configuration.update-layout',
  updateRender: 'injoffice.command.print-configuration.update-render',
  replace: 'injoffice.command.print-configuration.replace',
} as const

export const INJOFFICE_PRINT_UI_COMMANDS = {
  open: 'injoffice.command.print.open',
} as const

export interface UniverPrintConfigurationCommandConfig {
  updateLayout?: boolean
  updateRender?: boolean
  replace?: boolean
  /** Route public commands and undo/redo restores through a collaboration
   * session while retaining this adapter's persistence-first undo sink. */
  commandTargetFactory?: (controller: PrintConfigurationCommandController) => PrintConfigurationCommandTarget
}

export interface UniverPrintUIHost {
  /** Mount or reveal a host-owned PrintWorkspace. */
  open(): Promise<boolean | void> | boolean | void
}

export interface UniverPrintUIConfig {
  command?: boolean
  menu?: boolean
  title?: string
  tooltip?: string
  order?: number
  /** Runtime menu visibility supplied by a feature composition registry. */
  menuHidden$?: Observable<boolean>
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === 'object' && !Array.isArray(value)
}

/** Register workbook-scoped persisted settings and undo mutations. Each
 * public update command can be omitted independently. No print or preview UI
 * is registered by this adapter. */
export function registerUniverPrintConfigurationCommands(
  univer: Univer,
  unitId: string,
  manager: PrintManager,
  persistence: PrintConfigurationPersistence,
  config: UniverPrintConfigurationCommandConfig = {},
): { controller: PrintConfigurationCommandController; dispose(): void } {
  if (typeof unitId !== 'string' || !unitId.trim() || unitId.length > 512) throw new TypeError('unitId must contain 1-512 characters')
  const injector = univer.__getInjector()
  const commandService = injector.get(ICommandService)
  const undoRedo = injector.get(IUndoRedoService)
  const controller = new PrintConfigurationCommandController(manager, persistence, {
    push(record) {
      undoRedo.pushUndoRedo({
        unitID: unitId,
        undoMutations: [{ id: INJOFFICE_PRINT_CONFIGURATION_COMMANDS.restore, params: { snapshot: record.before } }],
        redoMutations: [{ id: INJOFFICE_PRINT_CONFIGURATION_COMMANDS.restore, params: { snapshot: record.after } }],
      })
    },
  })
  const registrations: IDisposable[] = []
  const target: PrintConfigurationCommandTarget = config.commandTargetFactory?.(controller) ?? controller
  if (!target || ['updateLayout', 'updateRender', 'replace', 'restore'].some((name) => typeof target[name as keyof PrintConfigurationCommandTarget] !== 'function')) throw new TypeError('commandTargetFactory must return a print configuration command target')
  const register = (id: string, handler: (params: unknown) => Promise<boolean>) => registrations.push(commandService.registerCommand({
    id,
    type: id === INJOFFICE_PRINT_CONFIGURATION_COMMANDS.restore ? CommandType.MUTATION : CommandType.COMMAND,
    handler: async (_accessor, params: unknown = {}) => {
      try { return await handler(params) } catch { return false }
    },
  }))

  register(INJOFFICE_PRINT_CONFIGURATION_COMMANDS.restore, (params) =>
    Promise.resolve(target.restore(isRecord(params) ? params.snapshot : undefined)))
  if (config.updateLayout !== false) register(INJOFFICE_PRINT_CONFIGURATION_COMMANDS.updateLayout, (params) =>
    Promise.resolve(target.updateLayout((isRecord(params) ? params.patch : undefined) as Partial<PrintLayoutConfig>)))
  if (config.updateRender !== false) register(INJOFFICE_PRINT_CONFIGURATION_COMMANDS.updateRender, (params) =>
    Promise.resolve(target.updateRender((isRecord(params) ? params.patch : undefined) as Partial<PrintRenderConfig>)))
  if (config.replace !== false) register(INJOFFICE_PRINT_CONFIGURATION_COMMANDS.replace, (params) =>
    Promise.resolve(target.replace((isRecord(params) ? params.snapshot : undefined) as PrintConfigurationSnapshotV1)))

  return {
    controller,
    dispose() {
      for (const registration of registrations.reverse()) registration.dispose()
      if (target !== controller) target.dispose?.()
    },
  }
}

/** Register a public command and optional Start/Layout ribbon entry that asks
 * the host to mount the React print workspace. The host retains portal,
 * renderer, and application-shell ownership. */
export function registerUniverPrintUI(
  univer: Univer,
  host: UniverPrintUIHost,
  config: UniverPrintUIConfig = {},
): { dispose(): void } {
  if (!host || typeof host.open !== 'function') throw new TypeError('print UI host must expose open()')
  const injector = univer.__getInjector()
  const registrations: IDisposable[] = []
  if (config.command !== false) {
    const commandService = injector.get(ICommandService)
    registrations.push(commandService.registerCommand({
      id: INJOFFICE_PRINT_UI_COMMANDS.open,
      type: CommandType.COMMAND,
      handler: async () => {
        try { return await host.open() !== false } catch { return false }
      },
    }))
  }
  if (config.menu !== false && config.command !== false) {
    injector.get(IMenuManagerService).mergeMenu({
      [RibbonPosition.START]: {
        [RibbonStartGroup.LAYOUT]: {
          [INJOFFICE_PRINT_UI_COMMANDS.open]: {
            order: config.order ?? 90,
            menuItemFactory: () => ({
              id: INJOFFICE_PRINT_UI_COMMANDS.open,
              title: config.title ?? 'Print',
              tooltip: config.tooltip ?? 'Preview and print workbook pages',
              type: MenuItemType.BUTTON,
              hidden$: config.menuHidden$,
            }),
          },
        },
      },
    })
  }
  return {
    dispose() { for (const registration of registrations.reverse()) registration.dispose() },
  }
}
