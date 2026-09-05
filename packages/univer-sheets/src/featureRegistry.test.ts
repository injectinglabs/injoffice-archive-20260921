import { describe, expect, it, vi } from 'vitest'
import { InjOfficeFeatureRegistry, type InjOfficeFeatureProvider } from './featureRegistry'

function provider(log: string[], key: string, options: Partial<InjOfficeFeatureProvider> = {}): InjOfficeFeatureProvider {
  return {
    ...options,
    load: vi.fn(async () => ({
      activate: async ({ visible }: { visible: boolean }) => {
        log.push(`activate:${key}:${visible}`)
        return {
          setVisible: async (next: boolean) => { log.push(`visible:${key}:${next}`) },
          dispose: async () => { log.push(`dispose:${key}`) },
        }
      },
    })),
  }
}

describe('InjOfficeFeatureRegistry', () => {
  it('lazy-loads only enabled features and activates dependencies first', async () => {
    const log: string[] = []
    const charts = provider(log, 'charts')
    const print = provider(log, 'print', { dependencies: ['charts'] })
    const registry = new InjOfficeFeatureRegistry({ charts, print })
    await registry.configure({ print: { enabled: true, visible: false } })
    expect(charts.load).toHaveBeenCalledOnce()
    expect(print.load).toHaveBeenCalledOnce()
    expect(log).toEqual(['activate:charts:true', 'activate:print:false'])
    expect(registry.snapshot().filter((entry) => entry.state === 'active').map((entry) => entry.key)).toEqual(['charts', 'print'])
  })

  it('rejects explicit dependency conflicts and missing providers before loading', async () => {
    const log: string[] = []
    const print = provider(log, 'print', { dependencies: ['charts'] })
    const registry = new InjOfficeFeatureRegistry({ print })
    await expect(registry.configure({ print: true, charts: false })).rejects.toMatchObject({ code: 'DEPENDENCY_DISABLED', feature: 'print' })
    await expect(registry.configure({ print: true })).rejects.toMatchObject({ code: 'MISSING_LOADER', feature: 'charts' })
    expect(print.load).not.toHaveBeenCalled()
  })

  it('negotiates required server capabilities before loading', async () => {
    const log: string[] = []
    const collaboration = provider(log, 'collaboration', { requiredServerCapabilities: ['collaboration.v1'] })
    await expect(new InjOfficeFeatureRegistry({ collaboration }).configure({ collaboration: true })).rejects.toMatchObject({ code: 'SERVER_CAPABILITY_MISSING' })
    const registry = new InjOfficeFeatureRegistry({ collaboration }, { serverCapabilities: ['collaboration.v1'] })
    await registry.configure({ collaboration: true })
    expect(collaboration.load).toHaveBeenCalledOnce()
  })

  it('hot-updates visibility and unloads disabled features', async () => {
    const log: string[] = []
    const registry = new InjOfficeFeatureRegistry({ charts: provider(log, 'charts') })
    await registry.configure({ charts: true })
    await registry.configure({ charts: { enabled: true, visible: false } })
    await registry.configure({ charts: false })
    expect(log).toEqual(['activate:charts:true', 'visible:charts:false', 'dispose:charts'])
  })

  it('refuses false visibility bookkeeping when an active adapter cannot hide its UI', async () => {
    const registry = new InjOfficeFeatureRegistry({
      charts: { load: async () => ({ activate: () => ({ dispose() {} }) }) },
    })
    await registry.configure({ charts: true })
    await expect(registry.configure({ charts: { enabled: true, visible: false } }))
      .rejects.toMatchObject({ code: 'VISIBILITY_UNSUPPORTED', feature: 'charts' })
    expect(registry.snapshot().find(({ key }) => key === 'charts')).toMatchObject({ state: 'active', visible: true })
  })

  it('isolates event observers from activation state', async () => {
    const log: string[] = []
    const registry = new InjOfficeFeatureRegistry({ charts: provider(log, 'charts') })
    registry.onEvent(() => { throw new Error('observer failed') })
    await expect(registry.configure({ charts: true })).resolves.toBeUndefined()
    expect(registry.snapshot().find(({ key }) => key === 'charts')?.state).toBe('active')
  })

  it('rolls back features newly activated by a failed transition', async () => {
    const log: string[] = []
    const registry = new InjOfficeFeatureRegistry({
      charts: provider(log, 'charts'),
      print: { dependencies: ['charts'], load: async () => { throw new Error('chunk unavailable') } },
    })
    await expect(registry.configure({ print: true })).rejects.toMatchObject({ code: 'ACTIVATION_FAILED', feature: 'print' })
    expect(log).toEqual(['activate:charts:true', 'dispose:charts'])
    expect(registry.snapshot().every((entry) => entry.state === 'disabled')).toBe(true)
  })

  it('serializes concurrent configurations and disposes in reverse activation order', async () => {
    const log: string[] = []
    const registry = new InjOfficeFeatureRegistry({
      charts: provider(log, 'charts'),
      print: provider(log, 'print', { dependencies: ['charts'] }),
    })
    await Promise.all([registry.configure({ print: true }), registry.configure({})])
    expect(log).toEqual(['activate:charts:true', 'activate:print:true', 'dispose:print', 'dispose:charts'])
  })
})
