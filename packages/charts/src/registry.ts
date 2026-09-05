// Tiny pub/sub between the ChartManager (which computes options from live
// cell data) and the ChartFloat components Univer mounts (which only render).
// Module-level on purpose: float-dom components are instantiated by Univer's
// component system, not by our tree, so props can't carry live references —
// only serializable data (the chartId). This registry is the meeting point.

type Listener = () => void

const options = new Map<string, Record<string, unknown>>()
const listeners = new Map<string, Set<Listener>>()

export function publishChartOption(chartId: string, option: Record<string, unknown>): void {
  options.set(chartId, option)
  listeners.get(chartId)?.forEach((l) => l())
}

export function getChartOption(chartId: string): Record<string, unknown> | undefined {
  return options.get(chartId)
}

export function dropChart(chartId: string): void {
  options.delete(chartId)
  listeners.delete(chartId)
}

export function subscribeChart(chartId: string, listener: Listener): () => void {
  let set = listeners.get(chartId)
  if (!set) {
    set = new Set()
    listeners.set(chartId, set)
  }
  set.add(listener)
  return () => {
    set.delete(listener)
  }
}
