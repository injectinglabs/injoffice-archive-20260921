import { describe, expect, it, vi } from 'vitest'
import { deferNestedReactRootStart, type DeferredStartScheduler } from './nestedReactRootLifecycle'

function controlledScheduler() {
  const pending = new Map<number, () => void>()
  let nextHandle = 0
  const scheduler: DeferredStartScheduler = {
    schedule(callback) {
      const handle = ++nextHandle
      pending.set(handle, callback)
      return handle
    },
    cancel(handle) {
      pending.delete(handle)
    },
  }
  return {
    scheduler,
    flush() {
      const callbacks = [...pending.values()]
      pending.clear()
      callbacks.forEach((callback) => callback())
    },
    pending,
  }
}

describe('nested React root lifecycle', () => {
  it('cancels the StrictMode probe before constructing a nested root', () => {
    const clock = controlledScheduler()
    const start = vi.fn()

    const stopProbe = deferNestedReactRootStart(start, clock.scheduler)
    stopProbe()
    clock.flush()

    expect(start).not.toHaveBeenCalled()
    expect(clock.pending.size).toBe(0)
  })

  it('starts the surviving effect once and disposes it once', () => {
    const clock = controlledScheduler()
    const dispose = vi.fn()
    const start = vi.fn(() => dispose)

    const stop = deferNestedReactRootStart(start, clock.scheduler)
    clock.flush()
    stop()
    stop()

    expect(start).toHaveBeenCalledOnce()
    expect(dispose).toHaveBeenCalledOnce()
  })

  it('models a StrictMode probe followed by the surviving effect', () => {
    const clock = controlledScheduler()
    const dispose = vi.fn()
    const start = vi.fn(() => dispose)

    const stopProbe = deferNestedReactRootStart(start, clock.scheduler)
    stopProbe()
    const stopLiveRoot = deferNestedReactRootStart(start, clock.scheduler)
    clock.flush()

    expect(start).toHaveBeenCalledOnce()
    stopLiveRoot()
    expect(dispose).toHaveBeenCalledOnce()
  })
})
