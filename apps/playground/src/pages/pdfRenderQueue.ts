/** One raster task at a time across pages and thumbnails. Cancelled queued work
 * never starts; running tasks receive the same signal through PDF.js. */
export function createPdfRenderQueue() {
  let tail: Promise<unknown> = Promise.resolve()
  return <T>(signal: AbortSignal, task: () => Promise<T>): Promise<T> => {
    const result = tail.catch(() => undefined).then(() => {
      signal.throwIfAborted()
      return task()
    })
    tail = result.catch(() => undefined)
    return result
  }
}

export function pdfPageWindow(page: number, count: number, radius = 1): number[] {
  return Array.from({ length: Math.max(0, Math.min(count, page + radius) - Math.max(1, page - radius) + 1) }, (_, index) => Math.max(1, page - radius) + index)
}
