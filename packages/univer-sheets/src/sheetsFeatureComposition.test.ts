import { beforeAll, describe, expect, it, vi } from 'vitest'

type CompositionModule = typeof import('./sheetsFeatureComposition') & typeof import('./univerFeatureProvider') & typeof import('./featureRegistry')
let subject: CompositionModule

beforeAll(async () => {
  vi.stubGlobal('Path2D', class Path2D {})
  subject = { ...await import('./sheetsFeatureComposition'), ...await import('./univerFeatureProvider'), ...await import('./featureRegistry') }
})

describe('InjOffice Sheets feature composition', () => {
  it('normalizes every OSS flag and advanced provider through one deterministic matrix', async () => {
    const activations: string[] = []
    const visibility = new Map<string, boolean[]>()
    const disposals: string[] = []
    const providers = Object.fromEntries(subject.INJOFFICE_FEATURE_KEYS.map((key) => [
      key,
      subject.createUniverFeatureProvider(({ menuHidden$ }) => {
        activations.push(key)
        const hidden: boolean[] = []
        visibility.set(key, hidden)
        const subscription = menuHidden$.subscribe((value) => hidden.push(value))
        return { dispose() { subscription.unsubscribe(); disposals.push(key) } }
      }),
    ]))
    const advanced = Object.fromEntries(subject.INJOFFICE_FEATURE_KEYS.map((key, index) => [
      key,
      { enabled: true, visible: index % 2 === 0 },
    ]))
    const composition = subject.createInjOfficeSheetsFeatureComposition({
      container: 'test',
      providers,
      features: { notes: false, threadComments: false, ...advanced },
    })

    const pluginNames = composition.bundle.presets.flatMap((preset) => preset.plugins.map((registration) =>
      (Array.isArray(registration) ? registration[0] : registration).pluginName))
    expect(pluginNames).not.toContain('SHEET_NOTE_PLUGIN')
    expect(pluginNames).not.toContain('SHEET_THREAD_COMMENT_BASE_PLUGIN')
    expect(pluginNames).toContain('SHEET_TABLE_PLUGIN')

    await composition.activate()
    expect(activations).toEqual(subject.INJOFFICE_FEATURE_KEYS)
    expect(composition.snapshot()).toHaveLength(subject.INJOFFICE_SHEETS_FEATURE_KEYS.length)
    expect(composition.snapshot().find(({ key }) => key === 'notes')).toMatchObject({ state: 'disabled', visible: false })
    expect(visibility.get('advancedFormula')).toEqual([false])
    expect(visibility.get('charts')).toEqual([true])

    await composition.configure({ charts: { enabled: true, visible: true } })
    expect(visibility.get('charts')).toEqual([true, false])
    expect(visibility.get('advancedFormula')).toEqual([false, true])
    expect(disposals).toEqual(expect.arrayContaining(subject.INJOFFICE_FEATURE_KEYS.filter((key) => key !== 'charts')))
    await expect(composition.configure({ notes: true })).rejects.toThrow('static Univer preset')
    await composition.dispose()
    expect(disposals.at(-1)).toBe('charts')
  })

  it('rolls back previously activated providers when a later real-style registrar fails', async () => {
    const dispose = vi.fn()
    const composition = subject.createInjOfficeSheetsFeatureComposition({
      container: 'test',
      features: { charts: true, pivots: true },
      providers: {
        charts: subject.createUniverFeatureProvider(() => ({ dispose })),
        pivots: subject.createUniverFeatureProvider(() => { throw new Error('Univer command registration failed') }),
      },
    })
    await expect(composition.activate()).rejects.toMatchObject({ code: 'ACTIVATION_FAILED', feature: 'pivots' })
    expect(dispose).toHaveBeenCalledOnce()
    expect(composition.snapshot().filter(({ source }) => source === 'injoffice-provider').every(({ state }) => state === 'disabled')).toBe(true)
  })

  it('rejects unknown keys and object options for inseparable Univer presets', () => {
    expect(() => subject.createInjOfficeSheetsFeatureComposition({ container: 'test', features: { mystery: true } as never })).toThrow('Unknown')
    expect(() => subject.createInjOfficeSheetsFeatureComposition({ container: 'test', features: { notes: { enabled: true, visible: false } } as never })).toThrow('must be a boolean')
  })
})
