import { BooleanNumber, HorizontalAlign, VerticalAlign, type ICellData, type IStyleData, type IWorkbookData } from '@univerjs/core'
import type { FileChartAnchor, FileChartInfo } from '@injoffice/charts'
import top100 from './top100Workbook.json'

export const EDITOR_OVERVIEW_SHEET_ID = 'sample-overview'
export const EDITOR_DATA_SHEET_ID = top100.sheetId
export const EDITOR_SAMPLE_DISCLOSURE = 'Illustrative sample values, not live market data. Source and as-of date are not provided.'

export const EDITOR_NUMBER_FORMATS = {
  marketCap: '$#,##0.0,,,"bn"',
  price: '$#,##0.00',
  change: '+0.00%;-0.00%;0.00%',
  billions: '$#,##0.0"bn"',
} as const

const header: IStyleData = { bl: BooleanNumber.TRUE, bg: { rgb: '#174A73' }, cl: { rgb: '#FFFFFF' } }
const numeric: IStyleData = { ht: HorizontalAlign.RIGHT }
const muted: IStyleData = { cl: { rgb: '#536579' }, fs: 10 }
const title: IStyleData = { bl: BooleanNumber.TRUE, fs: 17, cl: { rgb: '#174A73' } }

/** Fresh values and styles on each mount: edits must never mutate the seed. */
export function createEditorWorkbook(): Partial<IWorkbookData> {
  const data: Record<number, Record<number, ICellData>> = {}
  for (const [rowKey, cells] of Object.entries(top100.cellData)) {
    const row = Number(rowKey)
    data[row] = {}
    for (const [columnKey, original] of Object.entries(cells)) {
      const column = Number(columnKey)
      const cell: ICellData = { ...original }
      const style: IStyleData = row === 0 ? { ...header } : {
        ...(row % 2 === 0 ? { bg: { rgb: '#F1F6FA' } } : {}),
        ...(column === 0 || column >= 3 && column <= 5 ? numeric : {}),
      }
      if (row > 0 && column >= 3 && column <= 5) {
        style.n = { pattern: [EDITOR_NUMBER_FORMATS.marketCap, EDITOR_NUMBER_FORMATS.price, EDITOR_NUMBER_FORMATS.change][column - 3] }
      }
      // Number-format sections supply the sign; color supplements, never replaces it.
      if (row > 0 && column === 5 && typeof cell.v === 'number') {
        style.cl = { rgb: cell.v < 0 ? '#A42C36' : '#17613D' }
      }
      if (row >= 102) Object.assign(style, { bl: BooleanNumber.TRUE, bg: { rgb: '#E5EEF5' } })
      cell.s = style
      data[row][column] = cell
    }
  }
  data[0][1].v = 'Company'
  data[0][3].v = 'Sample cap (USD bn)'
  data[0][4].v = 'Sample price (USD)'
  data[0][5].v = 'Sample change'
  // Keep long summary labels in the wide company column rather than the ticker column.
  data[102][1] = { ...data[102][2], v: 'Total sample market cap' }
  data[103][1] = { ...data[103][2], v: 'Average sample price' }
  delete data[102][2]
  delete data[103][2]

  const overview: Record<number, Record<number, ICellData>> = {
    0: { 0: { v: 'Company watchlist', s: title } },
    1: { 0: { v: 'Illustrative sample • not live market data', s: muted } },
    2: { 0: { v: 'Source / as-of date: not provided', s: muted } },
    4: { 0: { v: 'Company', s: header }, 1: { v: 'USD billions', s: header } },
    11: { 0: { v: 'Try it', s: { bl: BooleanNumber.TRUE, cl: { rgb: '#174A73' } } } },
    12: { 0: { v: 'Edit a value in B6:B10.', s: muted } },
    13: { 0: { v: 'The chart follows your changes.', s: muted } },
    15: { 0: { v: 'Open “Sample data” for 100 rows, formatted prices and changes, and summary formulas.', s: muted } },
  }
  for (let index = 0; index < 5; index++) {
    const source = data[index + 1]
    overview[index + 5] = {
      0: { v: source[1].v, s: { bg: { rgb: index % 2 === 0 ? '#F1F6FA' : '#FFFFFF' } } },
      1: { v: Number(source[3].v) / 1_000_000_000, s: { ...numeric, n: { pattern: EDITOR_NUMBER_FORMATS.billions }, bg: { rgb: index % 2 === 0 ? '#F1F6FA' : '#FFFFFF' } } },
    }
  }
  const workbook: Partial<IWorkbookData> = {
    id: 'wb1',
    name: 'illustrative_company_watchlist.xlsx',
    sheetOrder: [EDITOR_OVERVIEW_SHEET_ID, EDITOR_DATA_SHEET_ID],
    defaultStyle: { ff: 'Arial', fs: 11, vt: VerticalAlign.MIDDLE, cl: { rgb: '#243649' }, pd: { l: 8, r: 8 } },
    sheets: {
      [EDITOR_OVERVIEW_SHEET_ID]: {
        id: EDITOR_OVERVIEW_SHEET_ID,
        name: 'Watchlist overview',
        cellData: overview,
        showGridlines: BooleanNumber.FALSE,
        defaultRowHeight: 25,
        rowData: { 0: { h: 38 } },
        columnData: { 0: { w: 155 }, 1: { w: 115 }, 2: { w: 16 }, 3: { w: 90 }, 4: { w: 90 }, 5: { w: 90 }, 6: { w: 90 } },
        mergeData: [0, 1, 2, 15].map(row => ({ startRow: row, endRow: row, startColumn: 0, endColumn: 6 })).concat(
          [12, 13].map(row => ({ startRow: row, endRow: row, startColumn: 0, endColumn: 1 })),
        ),
      },
      [EDITOR_DATA_SHEET_ID]: {
        id: EDITOR_DATA_SHEET_ID,
        name: 'Sample data',
        cellData: data,
        defaultRowHeight: 26,
        rowData: { 0: { h: 34 }, 102: { h: 30 }, 103: { h: 30 } },
        freeze: { xSplit: 0, ySplit: 1, startRow: 1, startColumn: -1 },
        columnData: { 0: { w: 56 }, 1: { w: 245 }, 2: { w: 100 }, 3: { w: 170 }, 4: { w: 160 }, 5: { w: 125 }, 6: { w: 150 } },
      },
    },
  }
  return structuredClone(workbook)
}

// A compact chart lives beside, never over, the editable overview table.
export const EDITOR_CHARTS: FileChartInfo[] = [{
  part: 'xl/charts/chart1.xml',
  type: 'column',
  title: 'Sample market cap (USD bn)',
  series: [{ categoriesRef: "'Watchlist overview'!$A$5:$A$10", valuesRef: "'Watchlist overview'!$B$5:$B$10" }],
}]

export const EDITOR_CHART_ANCHORS: Record<string, FileChartAnchor> = {
  'xl/charts/chart1.xml': { FromCol: 3, FromRow: 4, ToCol: 7, ToRow: 14 },
}
