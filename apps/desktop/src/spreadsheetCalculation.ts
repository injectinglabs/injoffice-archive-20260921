import type { NativeWorkbookV2 } from '@injoffice/sheets/browser';
import type { XlsxNativeCalculationCacheV1 } from '@injoffice/xlsx-wasm';
import type { LocalCalculationInput, LocalCalculationResult, LocalInputCell } from '../../../packages/formulas/src/localWorkbookCalculation';

/** Project source values only: saved formula caches are deliberately excluded. */
export function calculationInput(workbook: NativeWorkbookV2): LocalCalculationInput {
  // Formula groups can supply values in cells with no explicit formula element.
  // Until group expansion is modeled, none of those stored values may be inputs.
  if (workbook.sheets.some(sheet => sheet.cells.some(cell => cell.formula && cell.formula.type !== 'normal')) || workbook.unsupported.some(item => item.capability === 'formula-groups')) throw new Error('Local calculation cannot yet expand shared, array, or data-table formulas. Stored results remain unverified.');
  return { revision: workbook.revision, sheets: workbook.sheets.map(sheet => ({ id: sheet.id, name: sheet.name, cells: sheet.cells.map(cell => {
    const projected: LocalInputCell = { row: cell.row, column: cell.column };
    if (cell.formula) return { ...projected, formula: `=${cell.formula.text}` };
    const value = cell.value;
    if (!value) return projected;
    if (value.kind === 'number' && value.lexical !== undefined && Number.isFinite(Number(value.lexical))) projected.value = { kind: 'number', value: Number(value.lexical) };
    else if (value.kind === 'boolean' && /^(0|1|true|false)$/.test(value.lexical ?? '')) projected.value = { kind: 'boolean', value: value.lexical === '1' || value.lexical === 'true' };
    else if (value.kind === 'string') projected.value = { kind: 'string', value: value.text ?? '' };
    else if (value.kind === 'error' && value.lexical) projected.value = { kind: 'error', value: value.lexical };
    else projected.unsupported = `${cell.ref}: ${value.kind} storage is not qualified for local calculation.`;
    return projected;
  }) })) };
}
export function calculationCache(workbook: NativeWorkbookV2, result: LocalCalculationResult): XlsxNativeCalculationCacheV1 {
  if (result.sourceRevision !== workbook.revision) throw new Error('Calculation belongs to an older workbook revision.');
  return { engine: result.engine, cells: result.cells.map(cell => ({ sheet_id: cell.sheetId, cell: { row: cell.row, column: cell.column }, expected_formula: cell.formula, value: cell.status === 'calculated' || cell.status === 'error' ? cell.value ?? null : null })) };
}
export function createSpreadsheetCalculator({ timeoutMs = 30000 }: { timeoutMs?: number } = {}) {
  let worker: Worker | null = null;
  let pending: { reject(reason: Error): void; timer: ReturnType<typeof setTimeout> } | null = null;
  function terminate() { worker?.terminate(); worker = null; if (pending) { clearTimeout(pending.timer); pending.reject(new Error('Local calculation cancelled.')); pending = null; } }
  return {
    calculate(input: LocalCalculationInput): Promise<LocalCalculationResult> {
      if (pending) return Promise.reject(new Error('A local calculation is already running.'));
      worker ??= new Worker(new URL('./spreadsheetCalculation.worker.ts', import.meta.url), { type: 'module' });
      return new Promise((resolve, reject) => {
        const finish = () => { if (pending) clearTimeout(pending.timer); pending = null; };
        pending = { reject, timer: setTimeout(() => { finish(); worker?.terminate(); worker = null; reject(new Error('Local calculation exceeded 30 seconds. Stored results remain unverified.')); }, timeoutMs) };
        worker!.onmessage = event => {
          let error: unknown, result: unknown;
          try { const data = event.data; error = data && typeof data === 'object' ? data.error : undefined; result = data && typeof data === 'object' ? data.result : undefined; }
          catch (reason) { finish(); reject(reason instanceof Error ? reason : new Error(String(reason))); return; }
          finish();
          if (error) reject(new Error(String(error)));
          else if (result) resolve(result);
          else reject(new Error('Local calculation worker returned no result.'));
        };
        worker!.onerror = event => {
          const message = event.message || 'Local calculation worker failed.';
          finish();
          worker?.terminate();
          worker = null;
          reject(new Error(message));
        };
        worker!.postMessage(input);
      });
    }, terminate,
  };
}
