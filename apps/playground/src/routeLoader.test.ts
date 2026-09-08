import { describe, expect, it, vi } from 'vitest'
import { createRouteLoader } from './routeLoader'

function deferred() {
  let resolve!: () => void
  let reject!: (error: unknown) => void
  const promise = new Promise<void>((yes, no) => { resolve = yes; reject = no })
  return { promise, resolve, reject }
}

function setup() {
  const chunks = new Map<string, ReturnType<typeof deferred>>()
  const callbacks = {
    preload: vi.fn((route: string) => {
      const chunk = deferred()
      chunks.set(route, chunk)
      return chunk.promise
    }),
    pending: vi.fn(), ready: vi.fn(), failed: vi.fn(),
  }
  return { loader: createRouteLoader(callbacks), chunks, ...callbacks }
}

describe('route loader', () => {
  it('does not replace the current page before the requested chunk resolves', async () => {
    const state = setup()
    const request = state.loader.load('charts')
    expect(state.pending).toHaveBeenCalledWith('charts')
    expect(state.ready).not.toHaveBeenCalled()
    state.chunks.get('charts')!.resolve()
    await request
    expect(state.ready).toHaveBeenCalledExactlyOnceWith('charts')
  })

  it('ignores outdated successes and failures during rapid navigation', async () => {
    const state = setup()
    const first = state.loader.load('charts')
    const second = state.loader.load('pivots')
    const third = state.loader.load('overview')
    state.chunks.get('overview')!.resolve()
    await third
    state.chunks.get('pivots')!.reject(new Error('offline'))
    state.chunks.get('charts')!.resolve()
    await Promise.all([first, second])
    expect(state.ready).toHaveBeenCalledExactlyOnceWith('overview')
    expect(state.failed).not.toHaveBeenCalled()
  })

  it('handles a failed chunk and supports retrying the same route', async () => {
    const state = setup()
    const request = state.loader.load('charts')
    const error = new Error('offline')
    state.chunks.get('charts')!.reject(error)
    await request
    expect(state.failed).toHaveBeenCalledExactlyOnceWith('charts', error)
    expect(state.ready).not.toHaveBeenCalled()
    const retry = state.loader.load('charts')
    state.chunks.get('charts')!.resolve()
    await retry
    expect(state.ready).toHaveBeenCalledExactlyOnceWith('charts')
  })

  it('ignores completions after unmount', async () => {
    const state = setup()
    const request = state.loader.load('charts')
    state.loader.cancel()
    state.chunks.get('charts')!.resolve()
    await request
    expect(state.ready).not.toHaveBeenCalled()
    expect(state.failed).not.toHaveBeenCalled()
  })
})
