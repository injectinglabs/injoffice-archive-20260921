import type { NativeWorkbookV2, WorkbookMutationBatchV1 } from '@injoffice/sheets/browser';
import type { createXlsxWasmClient, adaptWorkbookMutationBatchV1 } from '@injoffice/xlsx-wasm';
import type { LocalCalculationResult } from '../../../packages/formulas/src/localWorkbookCalculation';
import { parseDelimited, encodeDelimited, DELIMITED_LIMITS, type DelimitedFormat } from './delimitedText.ts';

type Client = ReturnType<typeof createXlsxWasmClient>;
type ImportOptions = { signal?:AbortSignal; client?: Client; adapt?: typeof adaptWorkbookMutationBatchV1; onProgress?(completed: number, total: number): void };
/** Build a new XLSX exclusively with native mutations; no source text becomes a formula. */
export async function importDelimitedWorkbook(seed: Uint8Array, source: Uint8Array, format: DelimitedFormat, options: ImportOptions = {}): Promise<Uint8Array> {
  options.signal?.throwIfAborted();
  const rows = parseDelimited(source, format);
  if (rows.length * rows.reduce((width,row)=>Math.max(width,row.length),0) > DELIMITED_LIMITS.fields) throw new Error('The rectangular import extent exceeds 50,000 cells.');
  const module = await import('@injoffice/xlsx-wasm');
  const client = options.client ?? module.createXlsxWasmClient(), adapt = options.adapt ?? module.adaptWorkbookMutationBatchV1;
  let bytes: Uint8Array = seed.slice();
  try {
    let workbook = await client.extract(bytes,{signal:options.signal});
    if (workbook.sheets.length !== 1 || workbook.sheets[0].cells.some(cell => cell.formula || cell.value && (cell.value.kind !== 'string' || cell.value.text))) throw new Error('Delimited import requires a fresh blank workbook.');
    const apply = async (operations: WorkbookMutationBatchV1['operations']) => {
      const batch: WorkbookMutationBatchV1 = { protocol: 'injoffice.xlsx.mutations', version: 1, batch_id: crypto.randomUUID(), expected_revision: workbook.source.package_sha256, operations };
      bytes = await client.apply(bytes, workbook, adapt(workbook, batch),{signal:options.signal}); workbook = await client.extract(bytes,{signal:options.signal});
    };
    const sheetId = workbook.sheets[0].id;
    const seedRows = workbook.sheets[0].cells.reduce((count, cell) => Math.max(count, cell.row + 1), 0);
    if (seedRows > 10000) throw new Error('The blank workbook seed exceeds import bounds.');
    if (seedRows) await apply([{ operation_id: 'import-clear-seed', sheet_id: sheetId, kind: 'row.delete', index: 0, count: seedRows }]);
    const total = rows.reduce((count, row) => count + row.length, 0); let completed = 0, pending: WorkbookMutationBatchV1['operations'] = [], payloadBytes = 0;
    const flush = async () => { if (!pending.length) return; await apply(pending); completed += pending.length; options.onProgress?.(completed, total); pending = []; payloadBytes = 0; };
    for (let row = 0; row < rows.length; row++) for (let column = 0; column < rows[row].length; column++) {
      const operation = { operation_id: `import-${row}-${column}`, sheet_id: sheetId, kind: 'cell.set_value' as const, cell: { row, column }, value: rows[row][column] };
      const size = new TextEncoder().encode(JSON.stringify(operation)).byteLength;
      if (pending.length >= 2000 || payloadBytes + size > 1024 * 1024) await flush(); pending.push(operation); payloadBytes += size;
    }
    await flush();
    const cells = new Map(workbook.sheets[0].cells.map(cell => [`${cell.row}:${cell.column}`, cell]));
    for (let row = 0; row < rows.length; row++) for (let column = 0; column < rows[row].length; column++) {
      const cell = cells.get(`${row}:${column}`);
      if (!cell || cell.formula || cell.value && cell.value.kind !== 'string' || (cell.value?.text ?? '') !== rows[row][column]) throw new Error(`Imported text readback differs at row ${row + 1}, column ${column + 1}.`);
    }
    options.signal?.throwIfAborted();
    return bytes;
  } finally { if (!options.client) client.terminate(); }
}

export type DelimitedExportMode = 'values' | 'formulas';
export type VerifiedDelimitedCalculation = { revision: string; result: LocalCalculationResult };
/** Raw values exclude formatting; hidden rows are included, and stale formula caches never qualify. */
export function exportDelimitedSheet(workbook: NativeWorkbookV2, sheetId: string, format: DelimitedFormat, mode: DelimitedExportMode, calculation?: VerifiedDelimitedCalculation): Uint8Array {
  if (mode !== 'values' && mode !== 'formulas') throw new Error('Choose raw values or formula source.');
  const sheet = workbook.sheets.find(value => value.id === sheetId); if (!sheet) throw new Error('The selected worksheet is missing.');
  if (sheet.cells.some(cell => cell.formula && cell.formula.type !== 'normal') || workbook.unsupported.some(item => item.capability === 'formula-groups' && item.scope_id === `sheet:${sheetId}`)) throw new Error('Shared, array and data-table formula groups cannot yet be exported safely.');
  const height = sheet.cells.reduce((value, cell) => Math.max(value, cell.row + 1), 0), width = sheet.cells.reduce((value, cell) => Math.max(value, cell.column + 1), 0);
  if (height > DELIMITED_LIMITS.rows || width > DELIMITED_LIMITS.columns || height * width > DELIMITED_LIMITS.fields) throw new Error('The stored worksheet extent exceeds 10,000 rows, 256 columns or 50,000 fields.');
  const rows = Array.from({ length: height }, () => Array<string>(width).fill(''));
  const results = new Map(calculation?.result.cells.map(cell => [`${cell.sheetId}:${cell.row}:${cell.column}`, cell]) ?? []);
  for (const cell of sheet.cells) {
    let text = '';
    if (cell.formula) {
      if (mode === 'formulas') text = `=${cell.formula.text}`;
      else {
        const result = results.get(`${sheetId}:${cell.row}:${cell.column}`);
        if (calculation?.revision !== workbook.revision || !result || result.formula !== `=${cell.formula.text}` || result.status !== 'calculated' && result.status !== 'error' || !result.value) throw new Error(`Recalculate supported formulas before exporting raw values (${cell.ref}).`);
        text = result.value.kind === 'boolean' ? result.value.value ? 'TRUE' : 'FALSE' : String(result.value.value);
      }
    } else if (cell.value) text = cell.value.kind === 'string' ? cell.value.text ?? '' : cell.value.kind === 'boolean' ? cell.value.lexical === '1' || cell.value.lexical === 'true' ? 'TRUE' : 'FALSE' : cell.value.lexical ?? '';
    rows[cell.row][cell.column] = text;
  }
  return encodeDelimited(rows, format);
}
