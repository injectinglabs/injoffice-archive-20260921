import { describe, expect, it } from 'vitest'
import { assertCalculationResultCurrent } from '../../../packages/formulas/src/calculation'
import {
  createPlaygroundCalculator,
  evaluatePlaygroundFormula,
  playgroundCalculationRequest,
} from './formulaCalculation'

describe('playground formula calculation jobs', () => {
  it('evaluates arithmetic and reports protocol-bound results', async () => {
    expect(evaluatePlaygroundFormula('=2+3')).toEqual({ kind: 'value', value: 5 })
    expect(evaluatePlaygroundFormula('=SUM(4,6)')).toEqual({ kind: 'value', value: 10 })
    expect(evaluatePlaygroundFormula('=1/0')).toMatchObject({ kind: 'error', code: '#DIV/0!' })

    const { manager } = createPlaygroundCalculator()
    const result = await manager.submit(playgroundCalculationRequest('=2+3', 'job-add')).result
    expect(result.protocolVersion).toBe('injoffice.calculation.v1')
    expect(result.cells[0]?.result).toEqual({ kind: 'value', value: 5 })
    expect(result.engineFingerprint).toBe(manager.engineFingerprint)
    assertCalculationResultCurrent(result, {
      workbookId: 'playground-formulas',
      revision: 'rev-1',
      fingerprint: 'playground-formulas-rev-1',
    })
  })

  it('keeps audited fixture formulas as deterministic string values', async () => {
    const { manager } = createPlaygroundCalculator()
    const result = await manager.submit(playgroundCalculationRequest('=XLOOKUP(1,A1:A3,B1:B3)', 'job-lookup')).result
    expect(result.cells[0]?.result).toEqual({ kind: 'value', value: '=XLOOKUP(1,A1:A3,B1:B3)' })
  })
})
