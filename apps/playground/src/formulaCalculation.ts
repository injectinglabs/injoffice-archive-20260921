import {
  CALCULATION_PROTOCOL_VERSION,
  CalculationManager,
  type CalculationEngine,
  type CalculationEngineOutput,
  type FormulaCellValue,
} from '../../../packages/formulas/src/calculation'

export type PlaygroundWorkbook = { formulas: string[] }

const ADD = /^=(\d+(?:\.\d+)?)\s*\+\s*(\d+(?:\.\d+)?)$/
const SUM = /^=SUM\((\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\)$/i

export function evaluatePlaygroundFormula(formula: string): FormulaCellValue {
  if (formula === '=1/0') return { kind: 'error', code: '#DIV/0!', message: 'division by zero' }
  const add = ADD.exec(formula)
  if (add) return { kind: 'value', value: Number(add[1]) + Number(add[2]) }
  const sum = SUM.exec(formula)
  if (sum) return { kind: 'value', value: Number(sum[1]) + Number(sum[2]) }
  if (formula.startsWith('=')) return { kind: 'value', value: formula }
  return { kind: 'error', code: '#NAME?', message: 'formula must start with =' }
}

export function playgroundCalculationEngine(): CalculationEngine<PlaygroundWorkbook> {
  return {
    id: 'injoffice-playground',
    version: '1',
    deterministic: true,
    calculate(context): CalculationEngineOutput {
      return {
        cells: context.formulas.map((cell) => ({
          sheetId: cell.sheetId,
          row: cell.row,
          column: cell.column,
          result: evaluatePlaygroundFormula(cell.formula),
        })),
        diagnostics: [{ level: 'info', code: 'PLAYGROUND_ENGINE', message: 'Deterministic playground engine, not a workbook recalculation host.' }],
      }
    },
  }
}

export function createPlaygroundCalculator() {
  const engine = playgroundCalculationEngine()
  const manager = new CalculationManager(engine)
  return { engine, manager, protocolVersion: CALCULATION_PROTOCOL_VERSION }
}

export function playgroundCalculationRequest(formula: string, jobId: string) {
  return {
    jobId,
    workbookId: 'playground-formulas',
    sourceRevision: 'rev-1',
    sourceFingerprint: 'playground-formulas-rev-1',
    snapshot: { formulas: [formula] },
    formulas: [{ sheetId: 'sheet-1', row: 0, column: 0, formula }],
    metadata: { surface: 'formulas' },
  }
}
