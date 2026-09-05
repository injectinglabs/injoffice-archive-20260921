import { describe, expect, it } from 'vitest'
import { preprocessConnectorGrid, toRangeGrid } from './connectorPreprocess'

describe('playground connector range preprocessing', () => {
  it('trims text and parses numeric body cells without touching the header', async () => {
    const result = await preprocessConnectorGrid([
      [' account ', 'arr'],
      [' Northwind ', '128000'],
      [' Contoso ', '94,000'],
    ])
    expect(result.grid).toEqual([
      ['account', 'arr'],
      ['Northwind', 128000],
      ['Contoso', 94000],
    ])
    expect(result.stages.map((stage) => stage.id)).toEqual(['trim-text', 'parse-numbers'])
    expect(result.contract).toMatch(/^sha256:[0-9a-f]{64}$/)
    expect(result.fingerprint).toMatch(/^sha256:[0-9a-f]{64}$/)
    expect(result.raw[1]?.[0]).toBe(' Northwind ')
  })

  it('clones unknown cells into range values', () => {
    expect(toRangeGrid([['a', 1, null, true]])).toEqual([['a', 1, null, true]])
  })
})
