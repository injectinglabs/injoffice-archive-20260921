import { describe, expect, it } from 'vitest'
import type { IStyleData } from '@univerjs/core'
import { createEditorWorkbook, EDITOR_CHARTS, EDITOR_CHART_ANCHORS, EDITOR_DATA_SHEET_ID, EDITOR_NUMBER_FORMATS, EDITOR_OVERVIEW_SHEET_ID, EDITOR_SAMPLE_DISCLOSURE } from './editorWorkbook'

describe('editor workbook presentation', () => {
  it('opens a compact overview with an editable five-company chart and honest provenance', () => {
    const workbook = createEditorWorkbook()
    expect(workbook.sheetOrder).toEqual([EDITOR_OVERVIEW_SHEET_ID, EDITOR_DATA_SHEET_ID])
    const overview = workbook.sheets![EDITOR_OVERVIEW_SHEET_ID]
    expect(overview.cellData![1][0].v).toContain('not live market data')
    expect(overview.cellData![2][0].v).toContain('not provided')
    expect(EDITOR_SAMPLE_DISCLOSURE).toContain('Source and as-of date are not provided')
    expect(overview.cellData![5][0].v).toBe('NVIDIA')
    expect(overview.cellData![5][1].v).toBe(5331)
    expect(EDITOR_CHARTS[0].series[0].valuesRef).toBe("'Watchlist overview'!$B$5:$B$10")
    const anchor = EDITOR_CHART_ANCHORS['xl/charts/chart1.xml']
    expect(anchor.FromCol).toBeGreaterThan(1)
    expect(anchor.FromRow).toBe(4)
    expect(anchor.ToRow).toBeLessThanOrEqual(15)
    const rightEdge = Object.entries(overview.columnData!).filter(([column]) => Number(column) < anchor.ToCol).reduce((sum, [, data]) => sum + data.w!, 0)
    expect(rightEdge).toBeLessThanOrEqual(650)
  })

  it('preserves numeric values and formulas while applying number formats and a frozen header', () => {
    const data = createEditorWorkbook().sheets![EDITOR_DATA_SHEET_ID]
    expect(data.freeze).toEqual({ xSplit: 0, ySplit: 1, startRow: 1, startColumn: -1 })
    expect(data.cellData![0][3].v).toBe('Sample cap (USD bn)')
    expect(data.cellData![0][5].v).toBe('Sample change')
    expect(data.cellData![1][3]).toMatchObject({ v: 5331000000000, s: { n: { pattern: EDITOR_NUMBER_FORMATS.marketCap } } })
    expect(data.cellData![1][4]).toMatchObject({ v: 220.78, s: { n: { pattern: EDITOR_NUMBER_FORMATS.price } } })
    expect(data.cellData![16][5]).toMatchObject({ v: -0.004500000000000001, s: { n: { pattern: EDITOR_NUMBER_FORMATS.change }, cl: { rgb: '#A42C36' } } })
    expect(data.cellData![102][3].f).toBe('=SUM(D2:D101)')
    expect(data.cellData![103][4].f).toBe('=AVERAGE(E2:E101)')
    expect(data.cellData![100][1].v).toBe('T-Mobile US')
  })

  it('returns independent values and style objects for reset/remount', () => {
    const first = createEditorWorkbook()
    first.sheets![EDITOR_DATA_SHEET_ID].cellData![1][3].v = 1
    ;(first.sheets![EDITOR_OVERVIEW_SHEET_ID].cellData![4][0].s as IStyleData).bg!.rgb = '#000000'
    const second = createEditorWorkbook()
    expect(second.sheets![EDITOR_DATA_SHEET_ID].cellData![1][3].v).toBe(5331000000000)
    expect(second.sheets![EDITOR_OVERVIEW_SHEET_ID].cellData![5][1].v).toBe(5331)
    expect((second.sheets![EDITOR_OVERVIEW_SHEET_ID].cellData![4][0].s as IStyleData).bg!.rgb).toBe('#174A73')
  })
})
