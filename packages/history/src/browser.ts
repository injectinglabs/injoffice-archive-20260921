import { CommandType, ICommandService, type IDisposable, type Univer } from '@univerjs/core'
import { IMenuManagerService, MenuItemType, RibbonPosition, RibbonStartGroup } from '@univerjs/ui'
import type { Observable } from 'rxjs'
import {
  HistoryCommandController,
  isHistoryAuthor,
  type CaptureHistoryCommandOptions,
  type RestoreHistoryCommandOptions,
} from './commands'

export * from './index'

export const INJOFFICE_HISTORY_COMMANDS = {
  list: 'injoffice.command.history.list',
  preview: 'injoffice.command.history.preview',
  capture: 'injoffice.command.history.capture',
  restore: 'injoffice.command.history.restore',
  cancel: 'injoffice.command.history.cancel',
} as const

export interface UniverHistoryCommandConfig {
  list?: boolean
  preview?: boolean
  capture?: boolean
  restore?: boolean
  cancel?: boolean
}

/** Register configurable durable-history commands with Univer. The host keeps
 * ownership of storage, preview presentation, authorisation, and live reload. */
export function registerUniverHistoryCommands<TSnapshot>(
  univer: Univer,
  controller: HistoryCommandController<TSnapshot>,
  config: UniverHistoryCommandConfig = {},
): { dispose(): void } {
  const commandService = univer.__getInjector().get(ICommandService)
  const registrations: IDisposable[] = []
  const register = (enabled: boolean, id: string, handler: (params: any) => boolean | Promise<boolean>) => {
    if (!enabled) return
    registrations.push(commandService.registerCommand({
      id,
      type: CommandType.COMMAND,
      handler: (_accessor, params: any = {}) => handler(params),
    }))
  }

  register(config.list !== false, INJOFFICE_HISTORY_COMMANDS.list, async (params) => {
    if (!isObject(params) || (params.signal !== undefined && !isAbortSignal(params.signal))) return false
    await controller.list(params)
    return true
  })
  register(config.preview !== false, INJOFFICE_HISTORY_COMMANDS.preview, async (params) => {
    if (!isObject(params) || !isId(params.versionId)) return false
    await controller.preview(params.versionId, isAbortSignal(params.signal) ? params.signal : undefined)
    return true
  })
  register(config.capture !== false, INJOFFICE_HISTORY_COMMANDS.capture, async (params) => {
    if (!isObject(params) || !isHistoryAuthor(params.author)
      || (params.signal !== undefined && !isAbortSignal(params.signal))) return false
    await controller.capture(params as unknown as CaptureHistoryCommandOptions)
    return true
  })
  register(config.restore !== false, INJOFFICE_HISTORY_COMMANDS.restore, async (params) => {
    if (!isObject(params) || !isId(params.sourceVersionId) || !isHistoryAuthor(params.author)
      || (params.signal !== undefined && !isAbortSignal(params.signal))) return false
    await controller.restore(params as unknown as RestoreHistoryCommandOptions)
    return true
  })
  register(config.cancel !== false, INJOFFICE_HISTORY_COMMANDS.cancel, () => controller.cancel())

  return { dispose() { for (const registration of registrations.reverse()) registration.dispose() } }
}

export const INJOFFICE_HISTORY_UI_COMMANDS = {
  open: 'injoffice.command.history.open',
} as const

export interface UniverHistoryUIHost {
  /** Mount or reveal a host-owned HistoryTimeline sidebar. */
  open(): Promise<boolean | void> | boolean | void
}

export interface UniverHistoryUIConfig {
  command?: boolean
  menu?: boolean
  /** Call host.open() as soon as the adapter registers. Defaults to true. */
  autoOpen?: boolean
  title?: string
  tooltip?: string
  order?: number
  menuHidden$?: Observable<boolean>
}

/** Register a public command and optional Start/History ribbon item that asks
 * the host to mount HistoryTimeline. autoOpen mounts the sidebar immediately. */
export function registerUniverHistoryUI(
  univer: Univer,
  host: UniverHistoryUIHost,
  config: UniverHistoryUIConfig = {},
): { dispose(): void } {
  if (!host || typeof host.open !== 'function') throw new TypeError('history UI host must expose open()')
  const injector = univer.__getInjector()
  const registrations: IDisposable[] = []
  if (config.command !== false) {
    const commandService = injector.get(ICommandService)
    registrations.push(commandService.registerCommand({
      id: INJOFFICE_HISTORY_UI_COMMANDS.open,
      type: CommandType.COMMAND,
      handler: async () => {
        try { return await host.open() !== false } catch { return false }
      },
    }))
  }
  if (config.menu !== false && config.command !== false) {
    injector.get(IMenuManagerService).mergeMenu({
      [RibbonPosition.START]: {
        [RibbonStartGroup.HISTORY]: {
          [INJOFFICE_HISTORY_UI_COMMANDS.open]: {
            order: config.order ?? 40,
            menuItemFactory: () => ({
              id: INJOFFICE_HISTORY_UI_COMMANDS.open,
              title: config.title ?? 'Version history',
              tooltip: config.tooltip ?? 'Browse saved versions without changing the open workbook',
              type: MenuItemType.BUTTON,
              hidden$: config.menuHidden$,
            }),
          },
        },
      },
    })
  }
  if (config.autoOpen !== false) {
    try { void Promise.resolve(host.open()).catch(() => undefined) } catch { /* host mount is best-effort */ }
  }
  return {
    dispose() { for (const registration of registrations.reverse()) registration.dispose() },
  }
}

function isObject(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === 'object' && !Array.isArray(value)
}

function isId(value: unknown): value is string {
  return typeof value === 'string' && value.trim().length > 0 && value.length <= 512
}

function isAbortSignal(value: unknown): value is AbortSignal {
  return !!value && typeof value === 'object'
    && typeof (value as AbortSignal).aborted === 'boolean'
    && typeof (value as AbortSignal).addEventListener === 'function'
    && typeof (value as AbortSignal).removeEventListener === 'function'
}
