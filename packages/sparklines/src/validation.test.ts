import { describe, expect, it } from 'vitest'
import { validateSparkline, validateSparklineSnapshot } from './validation'
import type { SparklineSpec } from './types'

const valid: SparklineSpec = {
  id: 'spark-1', type: 'line',
  source: { sheetId: 's', startRow: 0, startColumn: 0, endRow: 0, endColumn: 4 },
  target: { sheetId: 's', row: 0, column: 5 },
}

describe('sparkline validation', () => {
  it('accepts the supported contract and rejects invalid bounds', () => {
    expect(validateSparkline(valid)).toEqual([])
    expect(validateSparkline({ ...valid, options: { min: 4, max: 2, lineWeight: 0 } }).map((issue) => issue.path)).toEqual([
      'options.lineWeight', 'options',
    ])
  })

  it('checks referential integrity in snapshots', () => {
    const issues = validateSparklineSnapshot({
      version: 1,
      sparklines: [{ ...valid, groupId: 'g' }],
      groups: [{ id: 'g', memberIds: ['spark-1', 'missing'] }],
    })
    expect(issues.map((issue) => issue.message)).toContain('unknown sparkline missing')
  })
})
