import { describe, expect, it } from 'vitest'
import { createPdfRenderQueue, pdfPageWindow } from './pdfRenderQueue'

describe('continuous PDF resource bounds', () => {
  it('bounds full-size and thumbnail windows at document edges', () => {
    expect(pdfPageWindow(1, 1000)).toEqual([1, 2])
    expect(pdfPageWindow(500, 1000)).toEqual([499, 500, 501])
    expect(pdfPageWindow(1000, 1000, 2)).toEqual([998, 999, 1000])
    expect(pdfPageWindow(1, 0)).toEqual([])
  })
  it('serializes work, skips aborted requests and survives failures', async () => {
    const queue = createPdfRenderQueue()
    const controller = new AbortController()
    const cancelled = new AbortController()
    let release!: () => void
    const seen: number[] = []
    const first = queue(controller.signal, async () => { seen.push(1); await new Promise<void>(resolve => { release = resolve }); seen.push(2) })
    const skipped = queue(cancelled.signal, async () => { seen.push(99) }).catch(() => undefined)
    const failed = queue(controller.signal, async () => { throw new Error('bad page') }).catch(() => undefined)
    const last = queue(controller.signal, async () => { seen.push(3) })
    cancelled.abort()
    await new Promise(resolve => setTimeout(resolve, 0))
    expect(seen).toEqual([1])
    release()
    await Promise.all([first, skipped, failed, last])
    expect(seen).toEqual([1, 2, 3])
  })
})
