import type { ComponentType } from 'react'
import { describe, expect, it } from 'vitest'
import { preloadableLazy, preloadableLazyNamed } from './preloadableLazy'

const Demo = (() => null) as ComponentType

describe('preloadable lazy routes', () => {
  it('shares one module request between repeated navigation intents', async () => {
    let loads = 0
    const route = preloadableLazy(async () => {
      loads += 1
      return { default: Demo }
    })

    await Promise.all([route.preload(), route.preload(), route.preload()])

    expect(loads).toBe(1)
  })

  it('preloads a named route export once', async () => {
    let loads = 0
    const route = preloadableLazyNamed(async () => {
      loads += 1
      return { Demo }
    }, 'Demo')

    await route.preload()
    await route.preload()

    expect(loads).toBe(1)
  })

  it('shares one failed request, then permits a later retry', async () => {
    let loads = 0
    const error = new Error('chunk unavailable')
    const route = preloadableLazy(async () => {
      loads += 1
      throw error
    })

    const first = route.preload()
    const sameAttempt = route.preload()
    await expect(first).rejects.toBe(error)
    await expect(sameAttempt).rejects.toBe(error)
    expect(loads).toBe(1)

    await expect(route.preload()).rejects.toBe(error)
    expect(loads).toBe(2)
  })
})
