export interface DeferredStartScheduler<Handle = number> {
  schedule(callback: () => void): Handle
  cancel(handle: Handle): void
}

const browserScheduler: DeferredStartScheduler = {
  schedule: (callback) => window.setTimeout(callback, 0),
  cancel: (handle) => window.clearTimeout(handle),
}

/**
 * Start a library-owned React root after the current outer React commit.
 *
 * React StrictMode probes effects with an immediate setup/cleanup/setup cycle.
 * Deferring the nested root lets that first cleanup cancel construction instead
 * of disposing a root whose own effects have not committed yet.
 */
export function deferNestedReactRootStart<Handle = number>(
  start: () => void | (() => void),
  scheduler: DeferredStartScheduler<Handle> = browserScheduler as DeferredStartScheduler<Handle>,
): () => void {
  let stopped = false
  let stop: void | (() => void)
  const handle = scheduler.schedule(() => {
    if (stopped) return
    stop = start()
  })

  return () => {
    if (stopped) return
    stopped = true
    scheduler.cancel(handle)
    stop?.()
  }
}
