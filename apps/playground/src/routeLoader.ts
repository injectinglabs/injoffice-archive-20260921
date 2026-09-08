/** Keep the current page usable until the latest requested route is ready. */
export function createRouteLoader<T>(callbacks: {
  preload: (route: T) => Promise<void>
  pending: (route: T) => void
  ready: (route: T) => void
  failed: (route: T, error: unknown) => void
}) {
  let generation = 0
  return {
    async load(route: T) {
      const attempt = ++generation
      callbacks.pending(route)
      try {
        await callbacks.preload(route)
        if (attempt === generation) callbacks.ready(route)
      } catch (error) {
        if (attempt === generation) callbacks.failed(route, error)
      }
    },
    cancel() { generation += 1 },
  }
}
