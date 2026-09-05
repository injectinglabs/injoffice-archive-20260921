import { describe, expect, it, vi } from 'vitest'
import {
  ClientFormulaError,
  ClientFormulaIntegration,
  createAdvancedFormulaFeatureProvider,
  type UniverFormulaFacadeLike,
} from './client'
import { TARGET_FUNCTIONS } from './target'

function facade(log: string[] = []): UniverFormulaFacadeLike {
  return {
    registerFunction: vi.fn((name) => {
      log.push(`register:${name}`)
      return { dispose: () => { log.push(`dispose:${name}`) } }
    }),
    registerAsyncFunction: vi.fn((name) => {
      log.push(`register-async:${name}`)
      return { dispose: () => { log.push(`dispose:${name}`) } }
    }),
    executeCalculation: vi.fn(() => { log.push('calculate') }),
    stopCalculation: vi.fn(() => { log.push('stop') }),
    onCalculationResultApplied: vi.fn(async () => { log.push('applied') }),
  }
}

describe('ClientFormulaIntegration', () => {
  it('registers normalized functions deterministically and unregisters in reverse order', () => {
    const log: string[] = []
    const integration = new ClientFormulaIntegration(facade(log), [
      { name: 'z_score', calculate: () => 1 },
      { name: ' acme.rate ', calculate: () => 2 },
      { name: 'middle', calculate: async () => 3, async: true },
    ])
    const activation = integration.activate()
    expect(activation.registeredFunctions).toEqual(['ACME.RATE', 'MIDDLE', 'Z_SCORE'])
    activation.dispose()
    activation.dispose()
    expect(log).toEqual([
      'register:ACME.RATE',
      'register-async:MIDDLE',
      'register:Z_SCORE',
      'dispose:Z_SCORE',
      'dispose:MIDDLE',
      'dispose:ACME.RATE',
    ])
  })

  it('validates the entire registration plan before touching the facade', () => {
    const host = facade()
    const integration = new ClientFormulaIntegration(host, [
      { name: 'PRIVATE_RATE', calculate: () => 1, requiredCapabilities: ['formula.private'] },
      { name: 'PUBLIC_RATE', calculate: () => 2 },
    ])
    expect(() => integration.activate()).toThrowError(expect.objectContaining({ code: 'CAPABILITY_MISSING', functionName: 'PRIVATE_RATE' }))
    expect(host.registerFunction).not.toHaveBeenCalled()
  })

  it('supports capability omission and an exact function allow-list', () => {
    const integration = new ClientFormulaIntegration(facade(), [
      { name: 'PUBLIC_RATE', calculate: () => 2 },
      { name: 'PRIVATE_RATE', calculate: () => 1, requiredCapabilities: ['formula.private'] },
      { name: 'UNSELECTED', calculate: () => 3 },
    ])
    const activation = integration.activate({
      enabledFunctions: ['private_rate', 'public_rate'],
      missingCapability: 'omit',
    })
    expect(activation.registeredFunctions).toEqual(['PUBLIC_RATE'])
    expect(activation.omittedFunctions).toEqual(['PRIVATE_RATE'])
  })

  it('rolls back prior registrations when a later registration fails', () => {
    const log: string[] = []
    const host = facade(log)
    vi.mocked(host.registerFunction)
      .mockImplementationOnce((name) => ({ dispose: () => { log.push(`rollback:${name}`) } }))
      .mockImplementationOnce(() => { throw new Error('host rejected') })
    const integration = new ClientFormulaIntegration(host, [
      { name: 'A_FN', calculate: () => 1 },
      { name: 'B_FN', calculate: () => 2 },
    ])
    expect(() => integration.activate()).toThrowError(expect.objectContaining({ code: 'REGISTRATION_FAILED', functionName: 'B_FN' }))
    expect(log).toEqual(['rollback:A_FN'])
  })

  it('refuses accidental replacement of an audited built-in target', () => {
    const integration = new ClientFormulaIntegration(facade(), [{ name: 'xlookup', calculate: () => 'replacement' }])
    expect(() => integration.activate()).toThrowError(expect.objectContaining({ code: 'TARGET_OVERRIDE_REFUSED', functionName: 'XLOOKUP' }))
    expect(integration.activate({ allowTargetOverride: true }).registeredFunctions).toEqual(['XLOOKUP'])
  })

  it('writes formulas through a worksheet range and controls calculation lifecycle', async () => {
    const log: string[] = []
    const integration = new ClientFormulaIntegration(facade(log))
    const events: string[] = []
    integration.onEvent((event) => events.push(event.type))
    const activation = integration.activate()
    const range = { setValue: vi.fn() }
    activation.setFormula(range, '=LET(x,2,x*3)')
    await activation.recalculate(5000)
    activation.stopCalculation()
    activation.dispose()

    expect(range.setValue).toHaveBeenCalledWith({ f: '=LET(x,2,x*3)' })
    expect(log).toEqual(['applied', 'calculate', 'stop'])
    expect(events).toEqual(['activated', 'calculation-started', 'calculation-completed', 'calculation-stopped', 'disposed'])
    expect(() => activation.setFormula(range, '=SUM(1,2)')).toThrowError(expect.objectContaining({ code: 'INACTIVE' }))
  })

  it('exposes the complete audited target vocabulary without duplicates', () => {
    const integration = new ClientFormulaIntegration(facade())
    expect(integration.targetVocabulary).toEqual(TARGET_FUNCTIONS.map(({ name }) => name))
    expect(new Set(integration.targetVocabulary).size).toBe(TARGET_FUNCTIONS.length)
    for (const { name } of TARGET_FUNCTIONS) expect(integration.supportsAuditedTarget(name.toLowerCase())).toBe(true)
  })

  it('adapts capability negotiation from the shared advancedFormula feature slot', async () => {
    const host = facade()
    const integration = new ClientFormulaIntegration(host, [
      { name: 'SECURE_RATE', calculate: () => 1, requiredCapabilities: ['formula.secure'] },
    ])
    const provider = createAdvancedFormulaFeatureProvider(integration)
    const loaded = await provider.load()
    const activation = loaded.activate({ serverCapabilities: new Set(['formula.secure']) })
    expect(provider.defaultEnabled).toBe(true)
    expect(activation.registeredFunctions).toEqual(['SECURE_RATE'])
  })

  it('keeps disabled configurations inert', () => {
    const host = facade()
    const activation = new ClientFormulaIntegration(host, [{ name: 'ACME_FN', calculate: () => 1 }]).activate({ enabled: false })
    expect(activation.enabled).toBe(false)
    expect(host.registerFunction).not.toHaveBeenCalled()
    expect(() => activation.stopCalculation()).toThrow(ClientFormulaError)
  })
})
