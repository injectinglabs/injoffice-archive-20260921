import { describe, expect, it } from 'vitest'
import { extractSparklineValues } from './extract'
import type { SparklineRangeRef } from './types'

const row: SparklineRangeRef = { sheetId: 's1', startRow: 2, endRow: 2, startColumn: 1, endColumn: 4 }
const column: SparklineRangeRef = { sheetId: 's1', startRow: 2, endRow: 5, startColumn: 1, endColumn: 1 }

describe('extractSparklineValues', () => {
  it('extracts rows with numbers, numeric text, percentages, and gaps', () => {
    expect(extractSparklineValues([[2, '3,400', '25%', 'nope', 999]], row)).toEqual([2, 3400, 0.25, null])
    expect(extractSparklineValues([[2, '', 4]], row, 'zero')).toEqual([2, 0, 4])
  })

  it('extracts a column in display order', () => {
    expect(extractSparklineValues([[1], [2], [null], [-4], [999]], column)).toEqual([1, 2, null, -4])
  })
})
