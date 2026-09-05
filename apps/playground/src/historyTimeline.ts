import { HistoryCommandController, type HistoryWorkspace } from '../../../packages/history/src/commands'
import type { HistoryManager } from '../../../packages/history/src/manager'
import type { HistoryPreview } from '../../../packages/history/src/types'
import type { HistorySnapshot } from './historyLifecycle'

export interface PlaygroundHistoryLive {
  snapshot: HistorySnapshot
  contentType: string
}

export function createPlaygroundHistoryWorkspace(options: {
  manager: HistoryManager<HistorySnapshot>
  live: () => PlaygroundHistoryLive
  onPreview: (preview: HistoryPreview<HistorySnapshot>) => void
  onActivate: (snapshot: HistorySnapshot) => void
}): HistoryWorkspace<HistorySnapshot> {
  return {
    captureSnapshot: async () => options.live(),
    showIsolatedPreview: async (preview) => {
      options.onPreview(preview)
    },
    activateRestoredVersion: async (version) => {
      const loaded = await options.manager.loadPreview(version.id)
      options.onActivate(loaded.snapshot)
    },
  }
}

export function createPlaygroundHistoryController(
  manager: HistoryManager<HistorySnapshot>,
  live: () => PlaygroundHistoryLive,
  onPreview: (preview: HistoryPreview<HistorySnapshot>) => void,
  onActivate: (snapshot: HistorySnapshot) => void,
): HistoryCommandController<HistorySnapshot> {
  return new HistoryCommandController(manager, createPlaygroundHistoryWorkspace({
    manager,
    live,
    onPreview,
    onActivate,
  }))
}
