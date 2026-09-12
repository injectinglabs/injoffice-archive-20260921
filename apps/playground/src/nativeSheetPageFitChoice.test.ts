import { describe, expect, it } from 'vitest'
import { nativeSheetPageFitChoice } from './nativeSheetPageFitChoice'

describe('explicit preview fit choices', () => {
  it('accepts one or both constrained dimensions including bounds', () => {
    expect(nativeSheetPageFitChoice('1', '1')).toEqual({ width: 1, height: 1 })
    expect(nativeSheetPageFitChoice('0', '100')).toEqual({ width: 0, height: 100 })
    expect(nativeSheetPageFitChoice('100', '0')).toEqual({ width: 100, height: 0 })
  })
  it('does not reinterpret missing, malformed or out-of-range inputs as unlimited', () => {
    for (const bad of ['', ' ', '-0', '-1', '1.5', 'NaN', 'Infinity', '1e1', '01', '101', '1000']) {
      expect(() => nativeSheetPageFitChoice(bad, '1')).toThrow('whole-number page limits')
      expect(() => nativeSheetPageFitChoice('1', bad)).toThrow('whole-number page limits')
    }
    expect(() => nativeSheetPageFitChoice('0', '0')).toThrow('at least one positive')
  })
})
