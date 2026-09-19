import { formatNativeSheetCellDisplayV2, type NativeWorkbookV2, type NativeWorkbookSheetV2, type NativeWorkbookCellV2, type SupportedWorkbookMutation, type StyleDelta, type WorkbookMutationBatchV1 } from '@injoffice/sheets/browser';

export type Position = { row: number; column: number };
export type Selection = { anchor: Position; end: Position };
export type Range = Position & { end_row: number; end_column: number };
export type SheetOperation = SupportedWorkbookMutation extends infer T ? T extends SupportedWorkbookMutation ? Omit<T, 'operation_id' | 'sheet_id'> : never : never;
export const MAX_ROWS = 1048576, MAX_COLUMNS = 16384, MAX_SELECTION = 10000;
export function address(cell: Position): string {
  let column = cell.column + 1, name = '';
  while (column > 0) { name = String.fromCharCode(65 + (column - 1) % 26) + name; column = Math.floor((column - 1) / 26); }
  return `${name}${cell.row + 1}`;
}
export function parseAddress(value: string): Position {
  const match = /^\$?([a-z]{1,3})\$?([1-9]\d*)$/i.exec(value.trim());
  if (!match) throw new Error('Enter a cell address such as B12 or a range such as A1:C8.');
  const column = [...match[1]!.toUpperCase()].reduce((n, letter) => n * 26 + letter.charCodeAt(0) - 64, 0) - 1;
  const row = Number(match[2]) - 1;
  if (row >= MAX_ROWS || column >= MAX_COLUMNS) throw new Error('That address is outside the worksheet.');
  return { row, column };
}
export function parseSelection(value: string): Selection {
  const parts = value.split(':'); if (parts.length > 2) throw new Error('Enter one rectangular range.');
  const anchor = parseAddress(parts[0]!); return { anchor, end: parts[1] ? parseAddress(parts[1]) : anchor };
}
export function selectedRange(selection: Selection): Range {
  return { row: Math.min(selection.anchor.row, selection.end.row), column: Math.min(selection.anchor.column, selection.end.column), end_row: Math.max(selection.anchor.row, selection.end.row), end_column: Math.max(selection.anchor.column, selection.end.column) };
}
export function contains(range: Range, cell: Position): boolean { return cell.row >= range.row && cell.row <= range.end_row && cell.column >= range.column && cell.column <= range.end_column; }
export function rangeSize(range: Range): number { return (range.end_row - range.row + 1) * (range.end_column - range.column + 1); }
export function selectionLabel(selection: Selection): string { const range = selectedRange(selection); const start = address(range), end = address({ row: range.end_row, column: range.end_column }); return start === end ? start : `${start}:${end}`; }
export function cellText(cell?: NativeWorkbookCellV2): string {
  if (cell?.formula) return `=${cell.formula.text.replace(/^=/, '')}`;
  const value = cell?.value;
  if (value?.kind === 'boolean') return value.lexical === '1' || value.lexical === 'true' ? 'TRUE' : 'FALSE';
  return value?.text ?? value?.lexical ?? '';
}
/** Preserve literal strings that would otherwise be interpreted as numbers/formulas on edit. */
export function editableCellText(cell?: NativeWorkbookCellV2): string {
  const text = cellText(cell);
  if (cell?.formula || cell?.value?.kind !== 'string' || !text) return text;
  const operation = valueOperation({ row: 0, column: 0 }, text);
  return operation.kind !== 'cell.set_value' || operation.value !== text ? `'${text}` : text;
}
export function cellDisplay(workbook: NativeWorkbookV2, cell?: NativeWorkbookCellV2): { text: string; note?: string } {
  if (!cell) return { text: '' };
  const value = cell.formula ? cell.formula.cached : cell.value;
  if (!value) return { text: cell.formula ? cellText(cell) : '', ...(cell.formula ? { note: 'Formula has no calculated value. Use Recalculate.' } : {}) };
  let text = value.text ?? value.lexical ?? '';
  if (value.kind === 'boolean') text = value.lexical === '1' || value.lexical === 'true' ? 'TRUE' : 'FALSE';
  if ((value.kind === 'number' || value.kind === 'date') && value.lexical !== undefined) {
    const format = workbook.styles.find(style => style.id === cell.style_id)?.effective.number_format;
    const display = formatNativeSheetCellDisplayV2(value.kind, value.lexical, format, workbook.date1904);
    if (display.status === 'ready') text = display.text;
    else return { text, note: `Stored value shown: ${display.message}` };
  }
  return { text, ...(cell.formula ? { note: 'Stored formula result; use Recalculate to verify it locally.' } : {}) };
}
export function valueOperation(cell: Position, input: string): SheetOperation {
  if (input === '') return { kind: 'cell.clear_value', cell };
  if (input.startsWith('=')) return { kind: 'cell.set_formula', cell, formula: input };
  let value: string | number | boolean = input;
  if (input.startsWith("'")) value = input.slice(1);
  else if (/^(true|false)$/i.test(input)) value = input.toLowerCase() === 'true';
  else if (/^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(input) && Number.isFinite(Number(input)) && input.replace(/[^0-9]/g, '').length <= 15) value = Number(input);
  return { kind: 'cell.set_value', cell, value };
}
/** TSV quoting matches spreadsheet clipboard conventions, including embedded line breaks. */
export function parsePaste(text: string): string[][] {
  if (text.length > 1_000_000) throw new Error('Paste at most 1 MB of text at once.');
  const rows: string[][] = [[]]; let field = '', quoted = false, closed = false;
  for (let i = 0; i < text.length; i++) {
    const char = text[i]!;
    if (quoted) { if (char === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else { quoted = false; closed = true; } } else field += char; continue; }
    if (char === '"' && field === '' && !closed) { quoted = true; continue; }
    if (char === '\t' || char === '\n' || char === '\r') { rows.at(-1)!.push(field); field = ''; closed = false; if (char !== '\t') { if (char === '\r' && text[i + 1] === '\n') i++; rows.push([]); } }
    else { if (closed) throw new Error('Invalid quoted clipboard text.'); field += char; }
    if (rows.length > MAX_SELECTION || rows.at(-1)!.length > MAX_SELECTION) throw new Error('Paste at most 10,000 cells at once.');
  }
  if (quoted) throw new Error('Clipboard text has an unclosed quote.');
  if (field !== '' || rows.at(-1)!.length || !/[\r\n]$/.test(text)) rows.at(-1)!.push(field); else rows.pop();
  const width = Math.max(...rows.map(row => row.length));
  if (rows.length * width > MAX_SELECTION) throw new Error('Paste at most 10,000 cells at once.');
  return rows.map(row => [...row, ...Array<string>(width - row.length).fill('')]);
}
export function pasteOperations(start: Position, text: string): { operations: SheetOperation[]; selection: Selection } {
  const rows = parsePaste(text), height = rows.length, width = rows[0]?.length ?? 0;
  if (!height || !width) throw new Error('Clipboard is empty.');
  if (start.row + height > MAX_ROWS || start.column + width > MAX_COLUMNS) throw new Error('Paste extends outside the worksheet.');
  return { operations: rows.flatMap((row, r) => row.map((value, c) => valueOperation({ row: start.row + r, column: start.column + c }, value))), selection: { anchor: start, end: { row: start.row + height - 1, column: start.column + width - 1 } } };
}
export function clearOperations(selection: Selection): SheetOperation[] {
  const range = selectedRange(selection); if (rangeSize(range) > MAX_SELECTION) throw new Error('Clear at most 10,000 cells at once.');
  return Array.from({ length: range.end_row - range.row + 1 }, (_, r) => Array.from({ length: range.end_column - range.column + 1 }, (_, c) => valueOperation({ row: range.row + r, column: range.column + c }, ''))).flat();
}
export function copySelection(sheet: NativeWorkbookSheetV2, selection: Selection): string {
  const range = selectedRange(selection); if (rangeSize(range) > MAX_SELECTION) throw new Error('Copy at most 10,000 cells at once.');
  const cells = new Map(sheet.cells.map(cell => [address(cell), cell]));
  return Array.from({ length: range.end_row - range.row + 1 }, (_, r) => Array.from({ length: range.end_column - range.column + 1 }, (_, c) => {
    const value = editableCellText(cells.get(address({ row: range.row + r, column: range.column + c })));
    return /[\t\r\n"]/.test(value) ? `"${value.replaceAll('"', '""')}"` : value;
  }).join('\t')).join('\r\n');
}
export function commandBatch(workbook: NativeWorkbookV2, sheetId: string, operations: SheetOperation[], id: string): WorkbookMutationBatchV1 {
  const sheet = workbook.sheets.find(value => value.id === sheetId);
  if (!sheet?.editable) throw new Error('This worksheet is read-only.');
  if (!operations.length || operations.length > MAX_SELECTION) throw new Error('Select between 1 and 10,000 cells.');
  return { protocol: 'injoffice.xlsx.mutations', version: 1, batch_id: id, expected_revision: workbook.source.package_sha256, operations: operations.map((operation, index) => ({ ...operation, sheet_id: sheetId, operation_id: `${id}-${index}` })) };
}

export type SpreadsheetRecoveryDraft = {
  version: 1; format: 'xlsx'; revision: string; sheetId: string; row: number; column: number; value: string;
};
export function recoveryDraft(workbook: NativeWorkbookV2, sheetId: string, position: Position, value: string): SpreadsheetRecoveryDraft {
  const result: SpreadsheetRecoveryDraft = { version: 1, format: 'xlsx', revision: workbook.source.package_sha256, sheetId, row: position.row, column: position.column, value };
  return validateRecoveryDraft(workbook, result);
}
/** Recovery restores pending text only onto the exact package and an editable native target. */
export function validateRecoveryDraft(workbook: NativeWorkbookV2, input: unknown): SpreadsheetRecoveryDraft {
  if (!input || typeof input !== 'object' || Array.isArray(input)) throw new Error('Invalid spreadsheet draft recovery record.');
  const draft = input as Record<string, unknown>;
  if (Object.keys(draft).sort().join(',') !== 'column,format,revision,row,sheetId,value,version' || draft.version !== 1 || draft.format !== 'xlsx' || draft.revision !== workbook.source.package_sha256 || typeof draft.sheetId !== 'string' || typeof draft.value !== 'string' || draft.value.length > 32767 || !Number.isInteger(draft.row) || !Number.isInteger(draft.column) || (draft.row as number) < 0 || (draft.row as number) >= MAX_ROWS || (draft.column as number) < 0 || (draft.column as number) >= MAX_COLUMNS) throw new Error('The recovered spreadsheet draft does not match this workbook.');
  const position = { row: draft.row as number, column: draft.column as number }, sheet = workbook.sheets.find(value => value.id === draft.sheetId);
  if (!sheet?.editable || sheet.state !== 'visible' || sheet.merged_ranges.some(range => contains(range, position)) || sheet.cells.find(cell => cell.row === position.row && cell.column === position.column)?.editable === false) throw new Error('The recovered cell is no longer editable.');
  return { version: 1, format: 'xlsx', revision: workbook.source.package_sha256, sheetId: sheet.id, ...position, value: draft.value };
}

/** Border presets change only requested edges; the native style table preserves all other components. */
export function borderOperations(selection: Selection, mode: 'all' | 'outer' | 'bottom' | 'clear', side: NonNullable<StyleDelta['border_top']>): SheetOperation[] {
 const range=selectedRange(selection), edge=mode==='clear'?{style:'none' as const,color:'#000000'}:{...side,color:side.color.toUpperCase()};
 if (mode==='all'||mode==='clear') return [{kind:'style.patch',range,style:{border_top:edge,border_bottom:edge,border_left:edge,border_right:edge}}];
 const bottom: SheetOperation={kind:'style.patch',range:{...range,row:range.end_row},style:{border_bottom:edge}};
 if(mode==='bottom')return [bottom];
 return [{kind:'style.patch',range:{...range,end_row:range.row},style:{border_top:edge}},bottom,{kind:'style.patch',range:{...range,end_column:range.column},style:{border_left:edge}},{kind:'style.patch',range:{...range,column:range.end_column},style:{border_right:edge}}];
}

/** Fill the viewport with visible rows even when a filter hides a long run. */
export function visibleRowWindow(start: number, count: number, hidden: ReadonlySet<number>, limit = MAX_ROWS): number[] {
  const rows: number[] = [];
  for (let row = Math.max(0, start); row < limit && rows.length < count; row++) if (!hidden.has(row)) rows.push(row);
  return rows;
}
export function visibleRowStep(start: number, delta: number, hidden: ReadonlySet<number>, limit = MAX_ROWS): number {
  if (!delta) return start;
  let row = Math.max(0, Math.min(limit - 1, start + delta));
  const direction = Math.sign(delta);
  while (row >= 0 && row < limit && hidden.has(row)) row += direction;
  return row < 0 || row >= limit ? start : row;
}
