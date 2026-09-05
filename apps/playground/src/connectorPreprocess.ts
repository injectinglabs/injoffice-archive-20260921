import { RangePreprocessPipeline, type RangeGrid, type RangeValue } from '../../../packages/connectors/src/preprocess'

function cloneGrid(grid: RangeGrid): RangeValue[][] {
  return grid.map((row) => [...row])
}

function asRangeValue(value: unknown): RangeValue {
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return value
  if (typeof value === 'number' && Number.isFinite(value)) return value
  if (value === undefined) return null
  return String(value)
}

export function toRangeGrid(grid: unknown[][]): RangeValue[][] {
  return grid.map((row) => row.map(asRangeValue))
}

export function playgroundConnectorPipeline(): RangePreprocessPipeline {
  const pipeline = new RangePreprocessPipeline()
  pipeline.register({
    id: 'trim-text',
    version: '1',
    order: 10,
    deterministic: true,
    process: (grid) => grid.map((row) => row.map((cell) => typeof cell === 'string' ? cell.trim() : cell)),
  })
  pipeline.register({
    id: 'parse-numbers',
    version: '1',
    order: 20,
    deterministic: true,
    process: (grid) => grid.map((row, rowIndex) => row.map((cell) => {
      if (rowIndex === 0 || typeof cell !== 'string') return cell
      const numeric = Number(cell.replace(/,/g, ''))
      return cell !== '' && Number.isFinite(numeric) ? numeric : cell
    })),
  })
  return pipeline
}

export async function preprocessConnectorGrid(grid: unknown[][], revision = 'playground-1') {
  const pipeline = playgroundConnectorPipeline()
  const result = await pipeline.run(toRangeGrid(grid), {
    range: { sheetId: 'connector-demo', startRow: 0, startColumn: 0, rowCount: Math.max(grid.length, 1), columnCount: Math.max(grid[0]?.length ?? 0, 1) },
    revision,
    mode: 'local',
  })
  return { ...result, raw: cloneGrid(toRangeGrid(grid)), contract: pipeline.fingerprint() }
}
