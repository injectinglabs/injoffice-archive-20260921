import { CommandType, ICommandService, IUndoRedoService, type IDisposable, type Univer } from '@univerjs/core'
import type { FUniver } from '@univerjs/core/lib/facade'
import type {} from '@univerjs/sheets/lib/facade'
import { IMenuManagerService, MenuItemType, RibbonDataGroup, RibbonPosition } from '@univerjs/ui'
import type { Observable } from 'rxjs'
import { ConnectorCommandController, type ConnectorCommandSnapshotV1 } from './commands'
import { ConnectorManager } from './manager'
import type { ConnectorRefreshOptions, ConnectorSpec, SourceFetcher } from './types'

export * from './index'

export const INJOFFICE_CONNECTOR_COMMANDS = {
  restore: 'injoffice.mutation.connector.restore',
  create: 'injoffice.command.connector.create',
  add: 'injoffice.command.connector.add',
  update: 'injoffice.command.connector.update',
  remove: 'injoffice.command.connector.remove',
  refresh: 'injoffice.command.connector.refresh',
} as const

export interface UniverConnectorCommandConfig {
  create?: boolean
  add?: boolean
  update?: boolean
  remove?: boolean
  refresh?: boolean
  menu?: boolean
  /** Runtime menu visibility supplied by a feature composition registry. */
  menuHidden$?: Observable<boolean>
  createMenu?: boolean
  refreshMenu?: boolean
  /** Supplies credential-free connector metadata for the Data ribbon. */
  createInput?: () => Omit<ConnectorSpec, 'id' | 'target'> | null
  /** Supplies the connector selected by host UI for the refresh button. */
  activeConnectorId?: () => string | null
}

/** Register connector commands through Univer's public command/undo services.
 * UI is omitted unless the host provides the corresponding picker callback. */
export function registerUniverConnectorCommands(
  univer: Univer,
  api: FUniver,
  fetchSource: SourceFetcher,
  config: UniverConnectorCommandConfig = {},
  manager?: ConnectorManager,
): { controller: ConnectorCommandController; dispose(): void } {
  const ownsManager = manager === undefined
  const connectorManager = manager ?? new ConnectorManager(api, fetchSource)
  const injector = univer.__getInjector()
  const commandService = injector.get(ICommandService)
  const undoRedo = injector.get(IUndoRedoService)
  const controller = new ConnectorCommandController(connectorManager, {
    push(record) {
      const unitID = api.getActiveWorkbook()?.getId()
      if (!unitID) return
      undoRedo.pushUndoRedo({
        unitID,
        undoMutations: [{ id: INJOFFICE_CONNECTOR_COMMANDS.restore, params: { snapshot: record.before } }],
        redoMutations: [{ id: INJOFFICE_CONNECTOR_COMMANDS.restore, params: { snapshot: record.after } }],
      })
    },
  })
  const registrations: IDisposable[] = []
  const register = (enabled: boolean, id: string, handler: (params: Record<string, unknown>) => boolean | Promise<boolean>) => {
    if (!enabled) return
    registrations.push(commandService.registerCommand({
      id,
      type: CommandType.COMMAND,
      handler: (_accessor, params: unknown = {}) => isObject(params) && handler(params),
    }))
  }
  registrations.push(commandService.registerCommand({
    id: INJOFFICE_CONNECTOR_COMMANDS.restore,
    type: CommandType.MUTATION,
    handler: (_accessor, params: unknown = {}) => isObject(params)
      && !!params.snapshot && controller.restore(params.snapshot as ConnectorCommandSnapshotV1),
  }))
  register(config.create !== false, INJOFFICE_CONNECTOR_COMMANDS.create, async (params) => {
    const input = isObject(params.input) ? params.input : config.createInput?.()
    if (!input) return false
    return await controller.createAtSelection(input as unknown as Omit<ConnectorSpec, 'id' | 'target'>) !== null
  })
  register(config.add !== false, INJOFFICE_CONNECTOR_COMMANDS.add, (params) =>
    isObject(params.spec) && controller.add(params.spec as unknown as ConnectorSpec) !== null)
  register(config.update !== false, INJOFFICE_CONNECTOR_COMMANDS.update, (params) =>
    isId(params.id) && isObject(params.patch)
      && controller.update(params.id, params.patch as Partial<Omit<ConnectorSpec, 'id'>>))
  register(config.remove !== false, INJOFFICE_CONNECTOR_COMMANDS.remove, (params) =>
    isId(params.id) && controller.remove(params.id))
  register(config.refresh !== false, INJOFFICE_CONNECTOR_COMMANDS.refresh, async (params) => {
    const id = isId(params.id) ? params.id : config.activeConnectorId?.()
    if (!id || !isRefreshOptions(params.options)) return false
    return await controller.refresh(id, (params.options ?? {}) as ConnectorRefreshOptions) === 'applied'
  })

  const showCreate = config.menu !== false && config.createMenu !== false
    && config.create !== false && !!config.createInput
  const showRefresh = config.menu !== false && config.refreshMenu !== false
    && config.refresh !== false && !!config.activeConnectorId
  if (showCreate || showRefresh) {
    const items: Record<string, unknown> = {}
    if (showCreate) items[INJOFFICE_CONNECTOR_COMMANDS.create] = {
      order: 40,
      menuItemFactory: () => ({
        id: INJOFFICE_CONNECTOR_COMMANDS.create,
        title: 'Data connection',
        tooltip: 'Create data connection',
        type: MenuItemType.BUTTON,
        hidden$: config.menuHidden$,
      }),
    }
    if (showRefresh) items[INJOFFICE_CONNECTOR_COMMANDS.refresh] = {
      order: 41,
      menuItemFactory: () => ({
        id: INJOFFICE_CONNECTOR_COMMANDS.refresh,
        title: 'Refresh connection',
        tooltip: 'Refresh selected data connection',
        type: MenuItemType.BUTTON,
        hidden$: config.menuHidden$,
      }),
    }
    injector.get(IMenuManagerService).mergeMenu({
      [RibbonPosition.DATA]: { [RibbonDataGroup.OTHERS]: items },
    })
  }

  return {
    controller,
    dispose() {
      for (const registration of registrations.reverse()) registration.dispose()
      if (ownsManager) connectorManager.dispose()
    },
  }
}

function isObject(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === 'object' && !Array.isArray(value)
}

function isId(value: unknown): value is string {
  return typeof value === 'string' && value.trim().length > 0 && value.length <= 512
}

function isRefreshOptions(value: unknown): value is ConnectorRefreshOptions | undefined {
  if (value === undefined) return true
  if (!isObject(value)) return false
  const allowed = new Set(['reason', 'signal', 'force', 'revision', 'mode', 'expectedPreprocessFingerprint'])
  if (Object.keys(value).some((key) => !allowed.has(key))) return false
  return (value.reason === undefined || value.reason === 'manual' || value.reason === 'onOpen' || value.reason === 'interval')
    && (value.signal === undefined || isAbortSignal(value.signal))
    && (value.force === undefined || typeof value.force === 'boolean')
    && (value.revision === undefined || typeof value.revision === 'string')
    && (value.mode === undefined || value.mode === 'local' || value.mode === 'collaborative')
    && (value.expectedPreprocessFingerprint === undefined || typeof value.expectedPreprocessFingerprint === 'string')
}

function isAbortSignal(value: unknown): value is AbortSignal {
  return !!value && typeof value === 'object'
    && typeof (value as AbortSignal).aborted === 'boolean'
    && typeof (value as AbortSignal).addEventListener === 'function'
    && typeof (value as AbortSignal).removeEventListener === 'function'
}
