import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it } from 'vitest'
import { SparklineCommandController, type SparklineUndoRecord } from './commands'
import { SparklineManager } from './manager'
import {
  SparklinePanel,
  SparklineSvg,
  createSparklinePanelCommandBindings,
  defaultSparklineInsertDraft,
  sparklineInputFromDraft,
} from './SparklinePanel'
import { compileSparklineGeometry } from './geometry'

function controller() {
  const records: SparklineUndoRecord[] = []
  const owner = new SparklineCommandController(new SparklineManager({ idFactory: (kind) => `${kind}-1` }), {
    push(record) { records.push(record) },
  })
  return { owner, records }
}

const source = { sheetId: 'sheet-1', startRow: 0, startColumn: 0, endRow: 0, endColumn: 2 }
const target = { sheetId: 'sheet-1', row: 1, column: 0 }

describe('SparklinePanel', () => {
  it('converts a labeled insert draft into a validated create payload', () => {
    const draft = defaultSparklineInsertDraft('s')
    draft.endColumn = '2'
    draft.targetRow = '1'
    draft.targetColumn = '0'
    draft.showHigh = true
    expect(sparklineInputFromDraft(draft)).toMatchObject({
      type: 'line',
      source: { sheetId: 's', startRow: 0, startColumn: 0, endRow: 0, endColumn: 2 },
      target: { sheetId: 's', row: 1, column: 0 },
      options: { emptyCells: 'gap', showHigh: true },
    })
    expect(() => sparklineInputFromDraft({ ...draft, endRow: '1' })).toThrow('must contain exactly one row or one column')
    expect(() => sparklineInputFromDraft({ ...draft, startRow: '-1' })).toThrow('startRow must be a non-negative integer')
  })

  it('routes insert, edit, group, and remove through snapshot undo', () => {
    const { owner, records } = controller()
    const commands = createSparklinePanelCommandBindings(owner)
    commands.create({ id: 'a', type: 'line', source, target })
    commands.create({
      id: 'b',
      type: 'line',
      source: { ...source, startRow: 1, endRow: 1 },
      target: { sheetId: 'sheet-1', row: 1, column: 3 },
    })
    records.length = 0
    expect(commands.update('a', { options: { showHigh: true } }).options).toEqual({ showHigh: true })
    expect(commands.group(['a', 'b']).memberIds).toEqual(['a', 'b'])
    commands.ungroup(['a'])
    expect(commands.remove('b')).toBe(true)
    expect(records.map(({ label }) => label)).toEqual([
      'Update sparkline',
      'Group sparklines',
      'Ungroup sparklines',
      'Remove sparkline',
    ])
  })

  it('renders an accessible insert form, cell preview, and group workflow', () => {
    const { owner } = controller()
    owner.create({ id: 'spark-a', type: 'column', source, target, options: { showHigh: true } })
    owner.create({
      id: 'spark-b',
      type: 'column',
      source: { ...source, startRow: 1, endRow: 1 },
      target: { sheetId: 'sheet-1', row: 2, column: 0 },
    })
    const markup = renderToStaticMarkup(
      <SparklinePanel
        controller={owner}
        readValues={() => [[1, 4, 2]]}
      />,
    )
    expect(markup).toContain('Insert sparkline')
    expect(markup).toContain('aria-label="Sparkline type"')
    expect(markup).toContain('aria-label="Source range"')
    expect(markup).toContain('aria-label="Target cell"')
    expect(markup).toContain('aria-label="Markers"')
    expect(markup).toContain('Group workflow')
    expect(markup).toContain('Group selected')
    expect(markup).toContain('Ungroup selected')
    expect(markup).toContain('column sparkline preview')
    expect(markup).toContain('spark-a')
    expect(markup).toContain('spark-b')
  })

  it('renders geometry through the host-mountable cell renderer', () => {
    const geometry = compileSparklineGeometry({
      type: 'line',
      values: [1, 3, 2],
      viewport: { width: 20, height: 10, padding: 1 },
    })
    const markup = renderToStaticMarkup(<SparklineSvg geometry={geometry} title="Revenue sparkline" />)
    expect(markup).toContain('role="img"')
    expect(markup).toContain('aria-label="Revenue sparkline"')
    expect(markup).toContain('<polyline')
  })
})
