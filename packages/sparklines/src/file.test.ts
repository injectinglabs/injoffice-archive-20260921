import { describe, expect, it } from 'vitest'
import { sparklinesFromFile } from './fromFile'
import { toWireSparklines } from './toFile'
import type { FileSparklineInfo } from './fromFile'
import type { SparklineSnapshotV1 } from './types'

const sheetIds: Record<string, string> = { Data: 'data-id', Dashboard: 'dash-id' }
const fromContext = { sheetIdOf: (name: string) => sheetIds[name] ?? null }
const toContext = { sheetNameOf: (id: string) => Object.entries(sheetIds).find(([, value]) => value === id)?.[0] ?? null }

function fileInfo(overrides: Partial<FileSparklineInfo> = {}): FileSparklineInfo {
  return {
    id: 'xl_worksheets_sheet1_xml_sparkline_1_1', groupId: 'native-group', type: 'line',
    sourceSheetName: 'Data', sourceRef: '$A$2:$E$2', targetSheetName: 'Dashboard', targetRef: '$F$2',
    options: { emptyCells: 'connect', showMarkers: true, min: 0, max: 10, colors: { series: '#112233' } },
    ...overrides,
  }
}

describe('native XLSX sparkline bridge', () => {
  it('hydrates and recomposes a native group with all three supported types', () => {
    const infos = [
      fileInfo(),
      fileInfo({ id: 'second', sourceRef: 'A3:E3', targetRef: 'F3' }),
      fileInfo({ id: 'column', groupId: undefined, type: 'column', sourceRef: 'A4:E4', targetRef: 'F4', options: { emptyCells: 'zero', showNegative: true } }),
      fileInfo({ id: 'win', groupId: undefined, type: 'win-loss', sourceRef: 'A5:E5', targetRef: 'F5', options: { rightToLeft: true } }),
    ]
    const loaded = sparklinesFromFile(infos, fromContext)
    expect(loaded.skipped).toEqual([])
    expect(loaded.snapshot.sparklines.map((item) => item.type)).toEqual(['line', 'line', 'column', 'win-loss'])
    expect(loaded.snapshot.groups).toHaveLength(1)
    expect(loaded.snapshot.groups[0].memberIds).toHaveLength(2)

    const wired = toWireSparklines(loaded.snapshot, toContext)
    expect(wired.skipped).toEqual([])
    expect(wired.groups).toHaveLength(3)
    expect(wired.groups[0]).toMatchObject({ targetSheetName: 'Dashboard', type: 'line', sparklines: [
      { sourceSheetName: 'Data', sourceRef: 'A2:E2', targetCellRef: 'F2' },
      { sourceSheetName: 'Data', sourceRef: 'A3:E3', targetCellRef: 'F3' },
    ] })
    expect(wired.groups.map((group) => group.type)).toEqual(['line', 'column', 'win-loss'])
  })

  it('skips a whole native group on diagnostics or an unsupported reference', () => {
    const diagnostics = sparklinesFromFile([
      fileInfo({ warnings: ['dateAxis is not representable'] }),
      fileInfo({ id: 'second', sourceRef: 'A3:E3', targetRef: 'F3' }),
    ], fromContext)
    expect(diagnostics.snapshot.sparklines).toEqual([])
    expect(diagnostics.skipped).toEqual(['sparkline native-group: dateAxis is not representable'])

    const reference = sparklinesFromFile([fileInfo({ sourceRef: 'A1:B2' })], fromContext)
    expect(reference.snapshot.sparklines).toEqual([])
    expect(reference.skipped[0]).toContain('A1 reference is unsupported')
  })

  it('refuses unrepresentable model groups without suppressing valid peers', () => {
    const snapshot: SparklineSnapshotV1 = {
      version: 1,
      sparklines: [
        { id: 'a', groupId: 'g', type: 'line', source: { sheetId: 'data-id', startRow: 0, startColumn: 0, endRow: 0, endColumn: 2 }, target: { sheetId: 'dash-id', row: 0, column: 3 }, options: { colors: { series: '#112233' } } },
        { id: 'b', groupId: 'g', type: 'line', source: { sheetId: 'data-id', startRow: 1, startColumn: 0, endRow: 1, endColumn: 2 }, target: { sheetId: 'data-id', row: 1, column: 3 }, options: { colors: { series: '#445566' } } },
        { id: 'c', type: 'column', source: { sheetId: 'data-id', startRow: 2, startColumn: 0, endRow: 2, endColumn: 2 }, target: { sheetId: 'dash-id', row: 2, column: 3 } },
      ],
      groups: [{ id: 'g', memberIds: ['a', 'b'] }],
    }
    const result = toWireSparklines(snapshot, toContext)
    expect(result.groups).toHaveLength(1)
    expect(result.groups[0].type).toBe('column')
    expect(result.skipped[0]).toContain('one worksheet')
  })

  it('reports invalid snapshots, missing sheets, and unsupported CSS colors', () => {
    const invalid = toWireSparklines({ version: 1, sparklines: [
      { id: '', type: 'line', source: { sheetId: 'data-id', startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }, target: { sheetId: 'dash-id', row: 0, column: 2 } },
    ], groups: [] }, toContext)
    expect(invalid.groups).toEqual([])
    expect(invalid.skipped[0]).toContain('id must not be empty')

    const missing = sparklinesFromFile([fileInfo({ sourceSheetName: 'Gone' })], fromContext)
    expect(missing.skipped[0]).toContain('sheet is unavailable')

    const badColor: SparklineSnapshotV1 = { version: 1, sparklines: [
      { id: 'one', type: 'line', source: { sheetId: 'data-id', startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }, target: { sheetId: 'dash-id', row: 0, column: 2 }, options: { colors: { series: 'red' } } },
    ], groups: [] }
    expect(toWireSparklines(badColor, toContext).skipped[0]).toContain('not a supported native RGB')
  })
})
