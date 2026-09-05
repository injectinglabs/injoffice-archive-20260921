import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import {
  createSheetsOutlineController,
  createSheetsSparklineController,
  outlineGroupFromRange,
  sparklineInputFromRange,
} from './sheetsPower'

describe('playground sheets power panels', () => {
  it('builds a one-dimensional sparkline insert from a selection', () => {
    expect(sparklineInputFromRange('line', 's1', { startRow: 1, startColumn: 0, endRow: 1, endColumn: 5 })).toMatchObject({
      type: 'line',
      source: { sheetId: 's1', startRow: 1, startColumn: 0, endRow: 1, endColumn: 5 },
      target: { sheetId: 's1', row: 1, column: 6 },
    })
    expect(sparklineInputFromRange('column', 's1', { startRow: 0, startColumn: 0, endRow: 2, endColumn: 2 })).toBeNull()
  })

  it('builds row and column outline groups from a selection', () => {
    const range = { startRow: 2, startColumn: 1, endRow: 5, endColumn: 3 }
    expect(outlineGroupFromRange('rows', 's1', 'row', range)).toMatchObject({ axis: 'row', start: 2, end: 5, collapsed: false })
    expect(outlineGroupFromRange('cols', 's1', 'column', range)).toMatchObject({ axis: 'column', start: 1, end: 3 })
  })

  it('creates controllers the Univer sidebar can mount', () => {
    const sparkline = createSheetsSparklineController()
    const created = sparkline.create(sparklineInputFromRange('line', 's1', { startRow: 0, startColumn: 0, endRow: 0, endColumn: 2 })!)
    expect(created.id).toContain('sparkline')
    const outline = createSheetsOutlineController()
    expect(outline.add(outlineGroupFromRange('q1', 's1', 'row', { startRow: 2, startColumn: 0, endRow: 4, endColumn: 0 })).ok).toBe(true)
    expect(outline.manager.list('s1', 'row')).toHaveLength(1)
  })

  it('mounts sparkline, outline, and shape panels in the sheets editor', () => {
    const editor = readFileSync(new URL('./UniverEditor.tsx', import.meta.url), 'utf8')
    const panels = readFileSync(new URL('./sheetsPowerPanels.tsx', import.meta.url), 'utf8')
    expect(editor).toContain('SheetsPowerPanels')
    expect(panels).toContain('SparklinePanel')
    expect(panels).toContain('OutlineGutter')
    expect(panels).toContain('ShapePanel')
    const tools = readFileSync(new URL('./pages/SheetsToolsPage.tsx', import.meta.url), 'utf8')
    expect(tools).toContain('PrintWorkspace')
  })
})
