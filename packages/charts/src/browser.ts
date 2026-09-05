import { CommandType, ICommandService, IUndoRedoService, type IDisposable, type Univer } from '@univerjs/core'
import type { FUniver } from '@univerjs/core/lib/facade'
import type {} from '@univerjs/sheets/lib/facade'
import { IMenuManagerService, MenuItemType, RibbonInsertGroup, RibbonPosition } from '@univerjs/ui'
import type { Observable } from 'rxjs'
import { ChartCommandController, type ChartSnapshotV1 } from './commands'
import type { FileChartAnchor } from './fromFile'
import { ChartImageExportManager, type ChartImageExportEvent, type ChartImageExportHost, type ChartImageExportOptions } from './imageExport'
import { ChartManager, type ChartLayerHost } from './manager'
import { isChartType, type ChartLayerOperation, type ChartSpec, type ChartType } from './types'

export * from './index'

export const INJOFFICE_CHART_COMMANDS = {
  restore: 'injoffice.mutation.chart.restore',
  create: 'injoffice.command.chart.create',
  add: 'injoffice.command.chart.add',
  update: 'injoffice.command.chart.update',
  remove: 'injoffice.command.chart.remove',
  layer: 'injoffice.command.chart.layer',
  exportImage: 'injoffice.command.chart.export-image',
} as const

export interface ChartCreateInput {
  type: ChartType
  title?: string
}

export interface UniverChartCommandConfig {
  create?: boolean
  add?: boolean
  update?: boolean
  remove?: boolean
  layer?: boolean
  exportImage?: boolean
  menu?: boolean
  /** Runtime menu visibility supplied by a feature composition registry. */
  menuHidden$?: Observable<boolean>
  /** Supplies chart type/title when the Insert ribbon invokes create. */
  createInput?: () => ChartCreateInput | null
  /** Mirrors back-to-front model order into the concrete drawing host. */
  layerHost?: ChartLayerHost
  /** Enables image export without coupling the package to one renderer. */
  imageExport?: {
    host: ChartImageExportHost
    timeoutMs?: number
    maxBytes?: number
    onEvent?: (event: ChartImageExportEvent) => void
  }
}

/** Register chart lifecycle commands against Univer's public command and undo
 * services. The Insert item is omitted unless the host supplies picker input. */
export function registerUniverChartCommands(
  univer: Univer,
  api: FUniver,
  config: UniverChartCommandConfig = {},
  manager?: ChartManager,
): { controller: ChartCommandController; exporter?: ChartImageExportManager; dispose(): void } {
  const injector = univer.__getInjector()
  const commandService = injector.get(ICommandService)
  const undoRedo = injector.get(IUndoRedoService)
  const chartManager = manager ?? new ChartManager(api, undefined, { layerHost: config.layerHost })
  const controller = new ChartCommandController(chartManager, {
    push(record) {
      const unitID = api.getActiveWorkbook()?.getId()
      if (!unitID) return
      undoRedo.pushUndoRedo({
        unitID,
        undoMutations: [{ id: INJOFFICE_CHART_COMMANDS.restore, params: { snapshot: record.before } }],
        redoMutations: [{ id: INJOFFICE_CHART_COMMANDS.restore, params: { snapshot: record.after } }],
      })
    },
  })
  const registrations: IDisposable[] = []
  registrations.push(commandService.registerCommand({
    id: INJOFFICE_CHART_COMMANDS.restore,
    type: CommandType.MUTATION,
    handler: (_accessor, params: { snapshot?: ChartSnapshotV1 } = {}) => !!params.snapshot && controller.restore(params.snapshot),
  }))
  if (config.create !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_CHART_COMMANDS.create,
    type: CommandType.COMMAND,
    handler: (_accessor, params: Partial<ChartCreateInput> = {}) => {
      const input = params.type ? params as ChartCreateInput : config.createInput?.()
      return !!input && isChartType(input.type)
        && (input.title === undefined || typeof input.title === 'string')
        && controller.createFromSelection(input.type, input.title) !== null
    },
  }))
  if (config.add !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_CHART_COMMANDS.add,
    type: CommandType.COMMAND,
    handler: (_accessor, params: { spec?: ChartSpec; cellAnchor?: FileChartAnchor } = {}) =>
      !!params.spec && controller.add(params.spec, params.cellAnchor) !== null,
  }))
  if (config.update !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_CHART_COMMANDS.update,
    type: CommandType.COMMAND,
    handler: (_accessor, params: { id?: string; patch?: Partial<Omit<ChartSpec, 'id' | 'nativeIdentity'>> } = {}) =>
      typeof params.id === 'string' && !!params.patch && controller.update(params.id, params.patch),
  }))
  if (config.remove !== false) registrations.push(commandService.registerCommand({
    id: INJOFFICE_CHART_COMMANDS.remove,
    type: CommandType.COMMAND,
    handler: (_accessor, params: { id?: string } = {}) => typeof params.id === 'string' && controller.remove(params.id),
  }))
  if (config.layer !== false && (config.layer === true || !!config.layerHost)) registrations.push(commandService.registerCommand({
    id: INJOFFICE_CHART_COMMANDS.layer,
    type: CommandType.COMMAND,
    handler: (_accessor, params: { id?: string; operation?: ChartLayerOperation } = {}) =>
      typeof params.id === 'string' && typeof params.operation === 'string' && controller.layer(params.id, params.operation),
  }))

  const exporter = config.imageExport && config.exportImage !== false
    ? new ChartImageExportManager((id) => chartManager.getSpec(id), config.imageExport.host, {
        timeoutMs: config.imageExport.timeoutMs,
        maxBytes: config.imageExport.maxBytes,
      })
    : undefined
  const disposeExportListener = exporter && config.imageExport?.onEvent ? exporter.onEvent(config.imageExport.onEvent) : undefined
  if (exporter) registrations.push(commandService.registerCommand({
    id: INJOFFICE_CHART_COMMANDS.exportImage,
    type: CommandType.COMMAND,
    handler: (_accessor, params: { id?: string; options?: ChartImageExportOptions; signal?: AbortSignal } = {}) =>
      typeof params.id === 'string' ? exporter.export(params.id, params.options, params.signal) : false,
  }))

  if (config.menu !== false && config.create !== false && config.createInput) {
    injector.get(IMenuManagerService).mergeMenu({
      [RibbonPosition.INSERT]: {
        [RibbonInsertGroup.MEDIA]: {
          [INJOFFICE_CHART_COMMANDS.create]: {
            order: 30,
            menuItemFactory: () => ({
              id: INJOFFICE_CHART_COMMANDS.create,
              title: 'Chart',
              tooltip: 'Insert chart',
              type: MenuItemType.BUTTON,
              hidden$: config.menuHidden$,
            }),
          },
        },
      },
    })
  }
  return {
    controller,
    exporter,
    dispose() {
      disposeExportListener?.()
      for (const registration of registrations.reverse()) registration.dispose()
    },
  }
}
