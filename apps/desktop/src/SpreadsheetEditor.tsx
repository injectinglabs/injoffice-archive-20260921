import { useEffect, useMemo, useRef, useState, type CSSProperties, type KeyboardEvent } from 'react';
import { createXlsxWasmClient, adaptWorkbookMutationBatchV1 } from '@injoffice/xlsx-wasm';
import type { XlsxNativeChart } from '@injoffice/xlsx-wasm';
import { SpreadsheetCharts } from './SpreadsheetCharts';
import { definedNameCaseKey } from '@injoffice/sheets/browser';
import type { NativeWorkbookV2, StyleDelta } from '@injoffice/sheets/browser';
import { visibleRowWindow, visibleRowStep, borderOperations, address, cellDisplay, editableCellText, clearOperations, commandBatch, contains, copySelection, MAX_COLUMNS, MAX_ROWS, parseSelection, pasteOperations, selectedRange, selectionLabel, valueOperation, recoveryDraft, validateRecoveryDraft, type SpreadsheetRecoveryDraft, type Position, type Selection, type SheetOperation } from './spreadsheetCommands';
import { calculationInput, calculationCache, createSpreadsheetCalculator } from './spreadsheetCalculation';
import type { LocalCalculationResult } from '../../../packages/formulas/src/localWorkbookCalculation';
import { exportDelimitedSheet, type DelimitedExportMode } from './spreadsheetDelimited';
import type { DelimitedFormat } from './delimitedText';
import { sheetLifecycleReason } from './spreadsheetSheetPolicy';
import './spreadsheet.css';

export interface OfficeEditorProps { registerHistory?: (commands: { undo(): void; redo(): void }) => void; registerCommit?: (commit: () => Promise<boolean>) => void; initialRecoveryDraft?: unknown; onRecoveryDraftChange?: (draft: unknown | null) => void; name: string; bytes: Uint8Array; onChange: (bytes: Uint8Array) => void; onBusyChange?: (busy: boolean) => void; onDraftChange?: (dirty: boolean) => void; viewOptions?: { zoom: number; navigation: boolean; focus: boolean } }

type Snapshot = { bytes: Uint8Array; workbook: NativeWorkbookV2; calculation?: LocalCalculationResult; calculationRevision?: string; charts: XlsxNativeChart[]; chartError?: string };
const initialSelection: Selection = { anchor: { row: 0, column: 0 }, end: { row: 0, column: 0 } };
const HISTORY_LIMIT = 20, HISTORY_BYTES = 128 * 1024 * 1024;
function historyPush(list: Snapshot[], entry: Snapshot): Snapshot[] {
  const result = [...list, entry].slice(-HISTORY_LIMIT); let bytes = result.reduce((sum, item) => sum + item.bytes.byteLength, 0);
  while (bytes > HISTORY_BYTES && result.length) bytes -= result.shift()!.bytes.byteLength;
  return result;
}

/** Local authoring over the native, revision-guarded XLSX transaction API. */
export function SpreadsheetEditor(props: OfficeEditorProps & { initialRecoveryDraft?: unknown; onRecoveryDraftChange?(draft: unknown | null): void; registerCommit?(commit: () => Promise<boolean>): void; registerHistory?(commands: { undo(): void; redo(): void }): void }) {
  const calculator = useRef<ReturnType<typeof createSpreadsheetCalculator> | null>(null);
  const historyLatest = useRef<(direction: 'undo' | 'redo') => void>(() => {});
  const callbacks = useRef(props); callbacks.current = props;
  const client = useRef<ReturnType<typeof createXlsxWasmClient> | null>(null);
  const current = useRef<Snapshot | null>(null), emitted = useRef<Uint8Array | null>(null), mounted = useRef(false), locked = useRef(false);
  const undo = useRef<Snapshot[]>([]), redo = useRef<Snapshot[]>([]);
  const [snapshot, setSnapshot] = useState<Snapshot | null>(null), [busy, setBusy] = useState(true), [error, setError] = useState('');
  const [sheetId, setSheetId] = useState(''), [selection, setSelection] = useState<Selection>(initialSelection), [origin, setOrigin] = useState({ row: 0, column: 0 });
  const draftTarget = useRef<SpreadsheetRecoveryDraft | null>(null), composing = useRef(false), commitLatest = useRef<() => Promise<boolean>>(async () => false);
  const [draft, setDraft] = useState<string | null>(null), draftRef = useRef<string | null>(null), [inline, setInline] = useState(false);
  const [borderLine, setBorderLine] = useState<NonNullable<StyleDelta['border_top']>['style']>('thin'), [borderColor, setBorderColor] = useState('#28764f');
  const [chartsOpen, setChartsOpen] = useState(false);
  const [namesOpen, setNamesOpen] = useState(false), [definedNameText, setDefinedNameText] = useState('');
  const [exportFormat, setExportFormat] = useState<DelimitedFormat>('csv'), [exportMode, setExportMode] = useState<DelimitedExportMode>('values');
  const [filterOpen, setFilterOpen] = useState(false), [filterColumn, setFilterColumn] = useState(0), [filterValues, setFilterValues] = useState(''), [filterBlank, setFilterBlank] = useState(false);
  const [sortOpen, setSortOpen] = useState(false), [sortColumn, setSortColumn] = useState(0), [sortHeader, setSortHeader] = useState(true), [sortDescending, setSortDescending] = useState(false);
  const [structureAction, setStructureAction] = useState<'row.insert' | 'row.delete' | 'column.insert' | 'column.delete' | null>(null);
  const [sheetAction, setSheetAction] = useState<'add' | 'rename' | 'delete' | null>(null), [sheetName, setSheetName] = useState('');
  const [location, setLocation] = useState('A1'), [rowHeight, setRowHeight] = useState('20'), [columnWidth, setColumnWidth] = useState('12');
  const [frozenHeights, setFrozenHeights] = useState<Record<number, number>>({});
  const [notice, setNotice] = useState(''), [, historyVersion] = useState(0);
  const grid = useRef<HTMLDivElement>(null), inlineInput = useRef<HTMLInputElement>(null), dragging = useRef(false);
  // Host busy state describes operations only; pending cell text uses onDraftChange.
  function setWorking(value: boolean) { locked.current = value; setBusy(value); callbacks.current.onBusyChange?.(value); }
  function updateDraft(value: string | null) {
    draftRef.current = value; setDraft(value); callbacks.current.onDraftChange?.(value !== null);
    if (value === null) { draftTarget.current = null; callbacks.current.onRecoveryDraftChange?.(null); }
    else if (current.current) {
      try { draftTarget.current = recoveryDraft(current.current.workbook, draftTarget.current?.sheetId ?? sheetId, draftTarget.current ?? selection.anchor, value); callbacks.current.onRecoveryDraftChange?.(draftTarget.current); }
      catch (reason) { setError(`Draft cannot be checkpointed: ${reason instanceof Error ? reason.message : String(reason)}`); }
    }
  }
  useEffect(() => {
    mounted.current = true; calculator.current = createSpreadsheetCalculator(); const worker = createXlsxWasmClient(); client.current = worker;
    return () => { mounted.current = false; calculator.current?.terminate(); calculator.current = null; worker.terminate(); client.current = null; callbacks.current.onBusyChange?.(false); callbacks.current.onDraftChange?.(false); };
  }, []);
  useEffect(() => {
    if (props.bytes === emitted.current) return;
    let cancelled = false; const worker = client.current!, pendingRecovery = props.initialRecoveryDraft; setWorking(true); setError(''); draftRef.current = null; draftTarget.current = null; setDraft(null); setInline(false); callbacks.current.onDraftChange?.(false); current.current = null; setSnapshot(null);
    worker.extract(props.bytes).then(async workbook => {
      const chartState = await readChartState(props.bytes, workbook);
      if (cancelled) return;
      const next = { bytes: props.bytes.slice(), workbook, ...chartState }; current.current = next; setSnapshot(next); undo.current = []; redo.current = [];
      setSheetId(workbook.sheets.find(sheet => sheet.state === 'visible')?.id ?? workbook.sheets[0]?.id ?? ''); setSelection(initialSelection); setOrigin({ row: 0, column: 0 }); setLocation('A1'); setNotice('');
      if (pendingRecovery != null) {
        try {
          const recovered = validateRecoveryDraft(workbook, pendingRecovery), position = { row: recovered.row, column: recovered.column };
          draftTarget.current = recovered; draftRef.current = recovered.value; setDraft(recovered.value); setInline(true); setSheetId(recovered.sheetId); setSelection({ anchor: position, end: position }); setLocation(address(position)); setOrigin({ row: Math.floor(position.row / 60) * 60, column: Math.floor(position.column / 20) * 20 });
          callbacks.current.onDraftChange?.(true); callbacks.current.onRecoveryDraftChange?.(recovered); setNotice('Recovered pending cell input. Apply or cancel to continue.');
        } catch (reason) { setError(reason instanceof Error ? reason.message : String(reason)); callbacks.current.onRecoveryDraftChange?.(null); }
      }
    }).catch(reason => { if (!cancelled) setError(String(reason instanceof Error ? reason.message : reason)); }).finally(() => { if (!cancelled) setWorking(false); });
    return () => { cancelled = true; };
  }, [props.bytes]);
  useEffect(() => { if (inline && draft !== null) inlineInput.current?.focus(); }, [inline, draft !== null]);
  const workbook = snapshot?.workbook, sheet = workbook?.sheets.find(value => value.id === sheetId);
  const calculatedCells = useMemo(() => new Map(snapshot?.calculation?.cells.map(cell => [`${cell.sheetId}:${cell.row}:${cell.column}`, cell]) ?? []), [snapshot?.calculation]);
  const cells = useMemo(() => new Map(sheet?.cells.map(cell => [address(cell), cell]) ?? []), [sheet]);
  const range = selectedRange(selection), active = cells.get(address(selection.anchor));
  const style = workbook?.styles.find(value => value.id === active?.style_id)?.effective;
  const merge = sheet?.merged_ranges.find(value => contains(value, selection.anchor));
  const canEdit = Boolean(sheet?.editable && !merge && active?.editable !== false);
  const selectedDefinedName = workbook?.defined_names?.find(name=>name.scope_sheet_id===undefined&&definedNameCaseKey(name.name)===definedNameCaseKey(definedNameText));
  function chooseDefinedName(name: NonNullable<NativeWorkbookV2['defined_names']>[number]) {
    if(locked.current||draftRef.current!==null)return;
    if(name.editable)setDefinedNameText(name.name);
    if(!name.target_sheet_id||!name.ref||!workbook?.sheets.some(sheet=>sheet.id===name.target_sheet_id&&sheet.state==='visible'))return;
    try{const next=parseSelection(name.ref);setSheetId(name.target_sheet_id);setSelection(next);setLocation(name.ref);setOrigin({row:next.anchor.row,column:next.anchor.column});}catch(reason){setError(String(reason));}
  }
  const disabled = busy || draft !== null || !sheet?.editable;
  const addSheetReason=sheetLifecycleReason(workbook,sheetId,'add'), deleteSheetReason=sheetLifecycleReason(workbook,sheetId,'delete');
  function reveal(position: Position) {
    setOrigin(previous => ({ row: position.row >= (sheet?.frozen_rows ?? 0) && (position.row < previous.row || position.row >= previous.row + 60) ? Math.floor(position.row / 60) * 60 : previous.row, column: position.column >= (sheet?.frozen_columns ?? 0) && (position.column < previous.column || position.column >= previous.column + 20) ? Math.floor(position.column / 20) * 20 : previous.column }));
  }
  function select(position: Position, extend = false) {
    if (locked.current || draftRef.current !== null) return;
    const merged = sheet?.merged_ranges.find(value => contains(value, position));
    if (merged && !extend) position = { row: merged.row, column: merged.column };
    const next = { anchor: extend ? selection.anchor : position, end: position }; setSelection(next); setLocation(selectionLabel(next)); reveal(position);
  }
  function startEdit(text = editableCellText(active), inCell = true) {
    if (locked.current || !canEdit) return;
    updateDraft(text); setInline(inCell); setError('');
  }
  async function readChartState(bytes: Uint8Array, workbook: NativeWorkbookV2): Promise<{charts:XlsxNativeChart[];chartError?:string}> {
    try { return { charts: (await client.current!.readCharts(bytes, workbook)).charts }; }
    catch (reason) { return { charts: [], chartError: reason instanceof Error ? reason.message : String(reason) }; }
  }
  async function calculateSnapshot(source: Snapshot): Promise<Snapshot> {
    const result = await calculator.current!.calculate(calculationInput(source.workbook));
    if (!result.cells.length) return source;
    const bytes = await client.current!.apply(source.bytes, source.workbook, { expected_revision: source.workbook.revision, calculation: calculationCache(source.workbook, result) });
    const workbook = await client.current!.extract(bytes);
    return { bytes, workbook, calculation: result, calculationRevision: workbook.revision, ...await readChartState(bytes, workbook) };
  }
  async function exportSheet() {
    const before = current.current;
    const bridge = window.injDesktop as typeof window.injDesktop & { exportBytes?(input: {name:string;bytes:Uint8Array}):Promise<{name:string}|null> };
    if (!before || locked.current || draftRef.current !== null || !bridge?.exportBytes) return;
    setWorking(true); setError('');
    try {
      const bytes = exportDelimitedSheet(before.workbook, sheetId, exportFormat, exportMode, before.calculation && before.calculationRevision ? {revision:before.calculationRevision,result:before.calculation} : undefined);
      const sheetName=before.workbook.sheets.find(value=>value.id===sheetId)?.name ?? 'Sheet';
      const result=await bridge.exportBytes({name:`${props.name.replace(/\.xlsx$/i,'')}-${sheetName}.${exportFormat}`,bytes});
      if (mounted.current && current.current===before && result) setNotice(`Exported ${result.name}; workbook unchanged`);
    } catch(reason) {if(mounted.current)setError(reason instanceof Error ? reason.message : String(reason));}
    finally {if(mounted.current&&current.current===before)setWorking(false);}
  }
  async function recalculate() {
    const before = current.current;
    if (!before || locked.current || draftRef.current !== null) return;
    if (!before.workbook.sheets.some(value => value.cells.some(cell => cell.formula))) { setNotice('This workbook has no formulas.'); return; }
    setWorking(true); setError('');
    try {
      const next = await calculateSnapshot(before);
      if (!mounted.current || current.current !== before) return;
      undo.current = historyPush(undo.current, before); redo.current = []; current.current = next; setSnapshot(next); emitted.current = next.bytes; callbacks.current.onChange(next.bytes);
      const unresolved = next.calculation?.cells.filter(cell => cell.status === 'unsupported' || cell.status === 'circular').length ?? 0;
      setNotice(unresolved ? `Calculated locally; ${unresolved} formulas remain unresolved (see cell details).` : 'Formulas calculated locally');
    } catch (reason) { if (mounted.current) setError(reason instanceof Error ? reason.message : String(reason)); }
    finally { if (mounted.current && (current.current === before || emitted.current === current.current?.bytes)) setWorking(false); }
  }
  async function execute(operations: SheetOperation[], message = 'Change applied', allowDraft = false): Promise<boolean> {
    const before = current.current, worker = client.current;
    if (!before || !worker || locked.current || (!allowDraft && draftRef.current !== null)) return false;
    setWorking(true); setError('');
    try {
      const batch = commandBatch(before.workbook, sheetId, operations, crypto.randomUUID());
      const bytes = await worker.apply(before.bytes, before.workbook, adaptWorkbookMutationBatchV1(before.workbook, batch));
      const nextWorkbook = await worker.extract(bytes);
      let next: Snapshot = { bytes, workbook: nextWorkbook, ...await readChartState(bytes, nextWorkbook) };
      if (next.workbook.sheets.some(value => value.cells.some(cell => cell.formula))) {
        try { next = await calculateSnapshot(next); } catch (reason) { setError(`Change saved; calculation is unavailable: ${reason instanceof Error ? reason.message : String(reason)}`); }
      }
      if (!mounted.current || current.current !== before) return false;
      undo.current = historyPush(undo.current, before); redo.current = []; current.current = next; setSnapshot(next); emitted.current = next.bytes;
      if (allowDraft) { updateDraft(null); setInline(false); }
      if (operations.some(operation => ['sheet.add','sheet.rename','sheet.delete'].includes(operation.kind))) {
        const nextSheet = operations[0]?.kind === 'sheet.add' ? next.workbook.sheets.at(-1) : next.workbook.sheets.find(value => value.id === sheetId) ?? next.workbook.sheets.find(value => value.state === 'visible');
        if (nextSheet) setSheetId(nextSheet.id); setSelection(initialSelection); setLocation('A1'); setOrigin({ row: 0, column: 0 }); setSheetAction(null);
      }
      if (operations.some(operation=>operation.kind==='sheet.filter')) {
        const filteredSheet=next.workbook.sheets.find(value=>value.id===sheetId);
        if (filteredSheet?.rows.some(row=>row.hidden&&row.row===selection.anchor.row)) {const start={row:range.row,column:range.column};setSelection({anchor:start,end:start});setOrigin({row:range.row,column:range.column});}
      }
      callbacks.current.onChange(next.bytes); setStructureAction(null); setSortOpen(false); setFilterOpen(false); setNotice(message); return true;
    } catch (reason) { if (mounted.current) setError(reason instanceof Error ? reason.message : String(reason)); return false; }
    finally { if (mounted.current && (current.current === before || emitted.current === current.current?.bytes)) setWorking(false); }
  }
  async function applyDraft(move?: Position): Promise<boolean> {
    if (draftRef.current === null || composing.current) return false;
    const formula = draftRef.current.startsWith('=');
    const applied = await execute([valueOperation(selection.anchor, draftRef.current)], formula ? 'Formula saved' : 'Cell updated', true);
    if (applied) {
      if (move) { const next = { anchor: move, end: move }; setSelection(next); setLocation(address(move)); reveal(move); }
      grid.current?.focus();
    }
    return applied;
  }
  commitLatest.current = async () => {
    if (!current.current || locked.current || composing.current) return false;
    return draftRef.current === null ? true : applyDraft();
  };
  useEffect(() => { let live = true; props.registerCommit?.(() => live ? commitLatest.current() : Promise.resolve(false)); return () => { live = false; }; }, [props.registerCommit]);
  function cancelDraft() { if (locked.current) return; updateDraft(null); setInline(false); grid.current?.focus(); }
  function restore(direction: 'undo' | 'redo') {
    if (locked.current || draftRef.current !== null || !current.current) return;
    const source = direction === 'undo' ? undo : redo, destination = direction === 'undo' ? redo : undo, next = source.current.pop();
    if (!next) return;
    destination.current = historyPush(destination.current, current.current); current.current = next; setSnapshot(next); emitted.current = next.bytes; if (!next.workbook.sheets.some(value => value.id === sheetId)) { setSheetId(next.workbook.sheets.find(value => value.state === 'visible')?.id ?? ''); setSelection(initialSelection); setLocation('A1'); setOrigin({ row: 0, column: 0 }); } callbacks.current.onChange(next.bytes); historyVersion(value => value + 1); setNotice(direction === 'undo' ? 'Change undone' : 'Change restored'); setError('');
  }
  historyLatest.current = restore;
  useEffect(() => { let live = true; props.registerHistory?.({ undo: () => { if (live) historyLatest.current('undo'); }, redo: () => { if (live) historyLatest.current('redo'); } }); return () => { live = false; }; }, [props.registerHistory]);
  function historyKeys(event: KeyboardEvent<HTMLElement>) {
    if (event.defaultPrevented || event.nativeEvent.isComposing || locked.current || draftRef.current !== null || event.altKey || !(event.metaKey || event.ctrlKey)) return;
    const target = event.target as HTMLElement; if (target.closest('input, textarea, [contenteditable]')) return;
    const key = event.key.toLowerCase(); if (key === 'z' || key === 'y') { event.preventDefault(); restore(event.shiftKey || key === 'y' ? 'redo' : 'undo'); }
  }
  function move(row: number, column: number, extend = false) { let base = extend ? selection.end : selection.anchor; if (!extend && merge) base = { row: row > 0 ? merge.end_row : merge.row, column: column > 0 ? merge.end_column : merge.column }; select({ row: visibleRowStep(base.row, row, hiddenRows), column: Math.min(MAX_COLUMNS - 1, Math.max(0, base.column + column)) }, extend); }
  function editKeys(event: KeyboardEvent<HTMLInputElement>) {
    if (event.nativeEvent.isComposing) return;
    if (event.key === 'Escape') { event.preventDefault(); cancelDraft(); }
    else if (event.key === 'Enter' || event.key === 'Tab') { event.preventDefault(); void applyDraft({ row: visibleRowStep(selection.anchor.row, event.key === 'Enter' ? event.shiftKey ? -1 : 1 : 0, hiddenRows), column: Math.min(MAX_COLUMNS - 1, Math.max(0, selection.anchor.column + (event.key === 'Tab' ? event.shiftKey ? -1 : 1 : 0))) }); }
  }
  function gridKeys(event: KeyboardEvent<HTMLDivElement>) {
    if (event.target !== event.currentTarget || locked.current || draftRef.current !== null || event.nativeEvent.isComposing) return;
    const command = event.metaKey || event.ctrlKey;
    if (command && (event.key.toLowerCase() === 'z' || event.key.toLowerCase() === 'y')) { event.preventDefault(); restore(event.shiftKey || event.key.toLowerCase() === 'y' ? 'redo' : 'undo'); return; }
    if (command || event.altKey) return;
    if (event.key === 'Home') { event.preventDefault(); select({ row: selection.anchor.row, column: 0 }, event.shiftKey); return; }
    const directions: Record<string, [number, number]> = { PageUp: [-60, 0], PageDown: [60, 0], ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1], Tab: [0, event.shiftKey ? -1 : 1] };
    if (directions[event.key]) { event.preventDefault(); move(...directions[event.key]!, event.key === 'Tab' ? false : event.shiftKey); }
    else if (event.key === 'Enter' || event.key === 'F2') { event.preventDefault(); startEdit(); }
    else if (event.key === 'Delete' || event.key === 'Backspace') { event.preventDefault(); try { void execute(clearOperations(selection), 'Selection cleared'); } catch (reason) { setError(String(reason)); } }
    else if (event.key.length === 1) { event.preventDefault(); startEdit(event.key); }
  }
  function format(patch: StyleDelta) { if (patch.font_color) patch={...patch,font_color:patch.font_color.toUpperCase()}; if (patch.fill_color) patch={...patch,fill_color:patch.fill_color.toUpperCase()}; void execute([{ kind: 'style.patch', range, style: patch }], 'Formatting applied'); }
  function applyStructure() {
    if (!structureAction) return;
    const rows = structureAction.startsWith('row.');
    const index = rows ? range.row : range.column, count = rows ? range.end_row - range.row + 1 : range.end_column - range.column + 1;
    void execute([{ kind: structureAction, index, count }], `${count} ${rows ? 'row(s)' : 'column(s)'} ${structureAction.endsWith('.insert') ? 'inserted' : 'deleted'}`);
  }
  function resize(kind: 'row' | 'column') {
    const value = Number(kind === 'row' ? rowHeight : columnWidth);
    if (!Number.isFinite(value) || value < 1 || value > (kind === 'row' ? 409.5 : 255)) { setError(kind === 'row' ? 'Row height must be 1–409.5 points.' : 'Column width must be 1–255 characters.'); return; }
    if ((kind === 'row' ? range.end_row - range.row + 1 : range.end_column - range.column + 1) > 10000) { setError('Resize at most 10,000 rows or columns.'); return; }
    const operations: SheetOperation[] = kind === 'row' ? Array.from({ length: range.end_row - range.row + 1 }, (_, offset) => ({ kind: 'row.set_height', row: range.row + offset, height_points: value })) : Array.from({ length: range.end_column - range.column + 1 }, (_, offset) => ({ kind: 'column.set_width', column: range.column + offset, width: value }));
    if (operations.length > 10000) { setError('Resize at most 10,000 rows or columns.'); return; } void execute(operations, `${kind === 'row' ? 'Row height' : 'Column width'} applied`);
  }
  const freezeSupported = (sheet?.frozen_rows ?? 0) <= 50 && (sheet?.frozen_columns ?? 0) <= 10;
  const frozenRows = freezeSupported ? sheet?.frozen_rows ?? 0 : 0, frozenColumns = freezeSupported ? sheet?.frozen_columns ?? 0 : 0;
  const hiddenRows = useMemo(() => new Set(sheet?.rows.filter(row=>row.hidden).map(row=>row.row) ?? []), [sheet]);
  const rows = [...new Set([...Array.from({ length: frozenRows }, (_, index) => index), ...visibleRowWindow(origin.row, 60, hiddenRows)])].filter(row => !hiddenRows.has(row));
  const columns = [...new Set([...Array.from({ length: frozenColumns }, (_, index) => index), ...Array.from({ length: Math.min(20, MAX_COLUMNS - origin.column) }, (_, index) => index + origin.column)])].filter(column => !sheet?.columns.find(value => column >= value.column && column <= value.end_column)?.hidden);
  function columnPixels(column: number) { const width = sheet?.columns.find(value => column >= value.column && column <= value.end_column)?.width ?? 12; return Math.max(24, Math.round(width * 7 + 5)); }
  useEffect(() => {
    if (!grid.current || !frozenRows || typeof ResizeObserver === 'undefined') return;
    const elements = [...grid.current.querySelectorAll<HTMLTableRowElement>('tr[data-sheet-row]')].filter(element => Number(element.dataset.sheetRow) < frozenRows);
    const measure = () => { const next = Object.fromEntries(elements.map(element => [Number(element.dataset.sheetRow), element.offsetHeight])); setFrozenHeights(previous => JSON.stringify(previous) === JSON.stringify(next) ? previous : next); };
    const observer = new ResizeObserver(measure); elements.forEach(element => observer.observe(element)); measure(); return () => observer.disconnect();
  }, [snapshot, sheetId, origin.row, origin.column, frozenRows]);
  function pinnedCellStyle(row: number, column: number): CSSProperties {
    const pinRow = row >= 0 && row < frozenRows, pinColumn = column >= 0 && column < frozenColumns;
    if (!pinRow && !pinColumn) return {};
    return { position: 'sticky', zIndex: pinRow && pinColumn ? 5 : 3,
      ...(pinRow ? { top: 27 + rows.filter(value => value < row).reduce((sum, value) => sum + (frozenHeights[value] ?? Math.max(27, (sheet?.rows.find(item => item.row === value)?.height_points ?? 20) * 96 / 72)), 0) } : {}),
      ...(pinColumn ? { left: 48 + columns.filter(value => value < column).reduce((sum,value) => sum + columnPixels(value), 0) } : {}), backgroundColor: '#fff', boxShadow: '1px 1px 0 #8da698' };
  }
  const activeStatus = merge ? 'Merged cell — unmerge to edit' : !canEdit ? 'Read-only cell' : draft !== null ? 'Editing — apply or cancel before saving' : 'Ready';
  return <section onKeyDown={historyKeys} className="sheet-editor" aria-label={`${props.name} spreadsheet`} aria-busy={busy}>
    <div className="sheet-tools" aria-label="Spreadsheet formatting">
      <div className="sheet-tool-group"><button disabled={disabled || !undo.current.length} onClick={() => restore('undo')}>Undo</button><button disabled={disabled || !redo.current.length} onClick={() => restore('redo')}>Redo</button><button disabled={disabled} onClick={() => void recalculate()}>Recalculate</button></div>
      <div className="sheet-tool-group">
        <select aria-label="Cell font" disabled={disabled} value={style?.font_name ?? 'Arial'} onChange={event => format({ font_name: event.target.value })}>{[...new Set(['Arial', 'Calibri', 'Times New Roman', 'Courier New', style?.font_name].filter(Boolean) as string[])].map(value => <option key={value}>{value}</option>)}</select>
        <select aria-label="Cell font size" disabled={disabled} value={style?.font_size_points ?? 11} onChange={event => format({ font_size_points: Number(event.target.value) })}>{[...new Set([8, 9, 10, 11, 12, 14, 16, 18, 24, 36, style?.font_size_points ?? 11])].sort((a,b) => a-b).map(value => <option key={value}>{value}</option>)}</select>
        <button className="sheet-bold" aria-label="Bold cells" aria-pressed={Boolean(style?.bold)} disabled={disabled} onClick={() => format({ bold: !style?.bold })}>B</button><button className="sheet-italic" aria-label="Italic cells" aria-pressed={Boolean(style?.italic)} disabled={disabled} onClick={() => format({ italic: !style?.italic })}>I</button>
        <label className="sheet-color">Text<input aria-label="Cell text color" type="color" disabled={disabled} value={style?.font_color ?? '#20242b'} onChange={event => format({ font_color: event.target.value })}/></label>
        <label className="sheet-color">Fill<input aria-label="Cell fill color" type="color" disabled={disabled} value={style?.fill_color ?? '#ffffff'} onChange={event => format({ fill_color: event.target.value })}/></label>
      </div>
      <div className="sheet-tool-group"><select aria-label="Cell alignment" disabled={disabled} value={style?.horizontal_alignment ?? 'general'} onChange={event => format({ horizontal_alignment: event.target.value as StyleDelta['horizontal_alignment'] })}><option value="general">Auto align</option><option value="left">Left</option><option value="center">Center</option><option value="right">Right</option></select><select aria-label="Cell vertical alignment" disabled={disabled} value={style?.vertical_alignment ?? 'bottom'} onChange={event => format({ vertical_alignment: event.target.value as StyleDelta['vertical_alignment'] })}><option value="top">Top</option><option value="middle">Middle</option><option value="bottom">Bottom</option></select><button aria-pressed={Boolean(style?.wrap_text)} disabled={disabled} onClick={() => format({ wrap_text: !style?.wrap_text })}>Wrap</button></div>
      <div className="sheet-tool-group"><select aria-label="Cell number format" disabled={disabled} value={style?.number_format ?? 'General'} onChange={event => format({ number_format: event.target.value })}>{[...new Set(['General', '0', '0.00', 'yyyy-mm-dd', style?.number_format ?? 'General'])].map(value => <option key={value} value={value}>{({ General: 'General', '0': 'Integer', '0.00': 'Decimal · 2 places', 'yyyy-mm-dd': 'Date · YYYY-MM-DD' } as Record<string,string>)[value] ?? `Existing: ${value}`}</option>)}</select><button disabled={disabled || Boolean(merge) || selectionLabel(selection) === address(selection.anchor)} onClick={() => void execute([{ kind: 'range.merge', range }], 'Cells merged; unmerge to edit their contents')}>Merge</button><button disabled={disabled || !merge} onClick={() => merge && void execute([{ kind: 'range.unmerge', range: { row: merge.row, column: merge.column, end_row: merge.end_row, end_column: merge.end_column } }], 'Cells unmerged')}>Unmerge</button></div>
<button disabled={!snapshot} aria-pressed={chartsOpen} onClick={() => setChartsOpen(value => !value)}>Charts</button>
<button disabled={busy || draft !== null} aria-pressed={namesOpen} onClick={()=>setNamesOpen(value=>!value)}>Names</button>
      <details className="sheet-dimensions sheet-export"><summary>Export sheet</summary><div><label>Format<select aria-label="Delimited export format" value={exportFormat} onChange={event=>setExportFormat(event.target.value as DelimitedFormat)}><option value="csv">CSV</option><option value="tsv">TSV</option></select></label><label>Content<select aria-label="Delimited export content" value={exportMode} onChange={event=>setExportMode(event.target.value as DelimitedExportMode)}><option value="values">Raw values · current calculation required</option><option value="formulas">Formula source · =expressions</option></select></label><small>UTF-8 with CRLF rows. Includes hidden rows and stored blank cells. Cell types, formatting and other sheets are not saved.</small><button disabled={busy || draft !== null || !sheet} onClick={()=>void exportSheet()}>Export selected sheet</button></div></details>
      <button disabled={disabled || range.row === range.end_row} onClick={() => { setFilterColumn(range.column); setFilterValues(''); setFilterBlank(false); setFilterOpen(true); }}>Filter text</button>
      <button disabled={disabled || !sheet?.auto_filter} onClick={() => void execute([{kind:'sheet.filter',filter:null}], 'Filter cleared; all filtered records shown')}>Clear filter</button>
      <button disabled={disabled || range.row === range.end_row} onClick={() => { setSortColumn(range.column); setSortOpen(true); }}>Sort range</button>
      <details className="sheet-dimensions"><summary>Freeze panes</summary><div><button disabled={disabled} onClick={() => void execute([{kind:'sheet.freeze',rows:1,columns:0}], 'Top row frozen')}>Freeze top row</button><button disabled={disabled} onClick={() => void execute([{kind:'sheet.freeze',rows:0,columns:1}], 'First column frozen')}>Freeze first column</button><button title="Freeze rows above and columns left of the active cell (up to 50 rows and 10 columns)" disabled={disabled || selection.anchor.row > 50 || selection.anchor.column > 10 || selection.anchor.row + selection.anchor.column === 0} onClick={() => void execute([{kind:'sheet.freeze',rows:selection.anchor.row,columns:selection.anchor.column}], 'Panes frozen at active cell')}>Freeze at selection</button><button disabled={disabled || !sheet?.frozen_rows && !sheet?.frozen_columns} onClick={() => void execute([{kind:'sheet.freeze',rows:0,columns:0}], 'Panes unfrozen')}>Unfreeze panes</button></div></details>
      <details className="sheet-dimensions"><summary>Rows & columns</summary><div>{(['row.insert','row.delete','column.insert','column.delete'] as const).map(action => <button key={action} disabled={disabled} onClick={() => setStructureAction(action)}>{action.endsWith('.insert') ? 'Insert' : 'Delete'} {action.startsWith('row.') ? 'rows' : 'columns'}</button>)}</div></details>
      <details className="sheet-dimensions"><summary>Cell size</summary><div><label>Row height (pt)<input aria-label="Row height in points" type="number" min="1" max="409.5" step="0.5" value={rowHeight} onChange={event => setRowHeight(event.target.value)}/></label><button disabled={disabled} onClick={() => resize('row')}>Set height</button><label>Column width (characters)<input aria-label="Column width in characters" type="number" min="1" max="255" step="0.5" value={columnWidth} onChange={event => setColumnWidth(event.target.value)}/></label><button disabled={disabled} onClick={() => resize('column')}>Set width</button></div></details>
    </div>
    <div className="sheet-formula-bar">
      <form onSubmit={event => { event.preventDefault(); try { const next = parseSelection(location); if (draftRef.current !== null || locked.current) return; setSelection(next); setLocation(selectionLabel(next)); reveal(next.anchor); grid.current?.focus(); } catch (reason) { setError(String(reason)); } }}><input aria-label="Cell or range address" disabled={busy || draft !== null} value={location} onChange={event => setLocation(event.target.value)}/></form>
      <span aria-hidden="true" className="sheet-fx">ƒx</span><input aria-label="Cell value or formula" maxLength={32767} onCompositionStart={() => { composing.current = true; }} onCompositionEnd={() => { composing.current = false; }} readOnly={busy || !canEdit} value={draft ?? editableCellText(active)} onChange={event => { updateDraft(event.target.value); setInline(false); }} onKeyDown={editKeys}/>
      {draft !== null && <><button disabled={busy} onClick={() => void applyDraft()}>Apply</button><button disabled={busy} onClick={cancelDraft}>Cancel</button></>}
    </div>
    {namesOpen && <section className="sheet-names" aria-label="Workbook names"><div className="sheet-names-heading"><strong>Workbook names</strong><button type="button" onClick={()=>setNamesOpen(false)}>Close</button></div>
      <p>Static named ranges persist in the workbook. Local formula calculation does not yet evaluate named references; their results remain unresolved.</p>
      <form onSubmit={event=>{event.preventDefault();if(sheetId)void execute([{kind:'name.set',name:definedNameText,target_sheet_id:sheetId,range}], 'Workbook name defined; formula caches invalidated');}}><label>Name<input aria-label="Defined name" maxLength={255} value={definedNameText} onChange={event=>setDefinedNameText(event.target.value)} placeholder="Quarterly_Total"/></label><span>Refers to {sheet?.name}!{selectionLabel(selection)}</span><button type="submit" disabled={disabled || !definedNameText || Boolean(selectedDefinedName&&!selectedDefinedName.editable)}>{selectedDefinedName?'Update range':'Define selection'}</button><button type="button" disabled={disabled || !selectedDefinedName?.editable} onClick={()=>void execute([{kind:'name.delete',name:selectedDefinedName!.name}], 'Unused workbook name deleted')}>Delete name</button></form>
      <div className="sheet-names-list">{workbook?.defined_names?.length?workbook.defined_names.map((name,index)=><div key={`${name.scope_sheet_id??'workbook'}:${name.name}:${index}`} className="sheet-name-entry"><button title={name.formula} disabled={busy || draft!==null || !name.target_sheet_id} onClick={()=>chooseDefinedName(name)}>{name.name}</button><code title={name.formula}>{name.formula}</code><span>{name.scope_sheet_id?`Sheet ${workbook.sheets.find(sheet=>sheet.id===name.scope_sheet_id)?.name??name.scope_sheet_id}`:'Workbook'}{name.editable?'':' · read-only'}</span></div>):<span>No defined names yet.</span>}</div>
    </section>}
    {sheet?.auto_filter && <p className="sheet-filter-status">Text filter on {sheet.auto_filter.ref}. Hidden records are preserved. Clear the filter before editing its text column or resizing its rows.</p>}
    {filterOpen && <form className="sheet-lifecycle" aria-label="Filter selected range" onSubmit={event => { event.preventDefault(); const values = filterValues.split('\n').filter(value => value.length > 0); void execute([{kind:'sheet.filter',filter:{ref:`${address({row:range.row,column:range.column})}:${address({row:range.end_row,column:range.end_column})}`,column:filterColumn,values,blank:filterBlank}}], 'Exact text filter applied'); }}>
      <span>Filter {selectionLabel(selection)} · first row is a header</span><label>Text column<select aria-label="Filter text column" value={filterColumn} onChange={event=>setFilterColumn(Number(event.target.value))}>{Array.from({length:Math.min(100,range.end_column-range.column+1)},(_,index)=>range.column+index).map(column=><option key={column} value={column}>{address({row:0,column}).replace(/\d+$/,'')}</option>)}</select></label>
      <label>Keep these exact values<textarea aria-label="Filter values, one per line" value={filterValues} onChange={event=>setFilterValues(event.target.value)} rows={3} placeholder="Green
Blue"/></label><label><input type="checkbox" checked={filterBlank} onChange={event=>setFilterBlank(event.target.checked)}/>Include blank cells</label>
      <small>Case-insensitive text matching. Numbers, dates, formula criteria, grouped rows and manual hidden rows are preserved and cannot be filtered here.</small>
      <button type="submit" disabled={disabled || (!filterValues.split('\n').some(value=>value.length) && !filterBlank) || filterColumn < range.column || filterColumn > range.end_column}>Apply filter</button><button type="button" disabled={busy} onClick={()=>setFilterOpen(false)}>Cancel</button>
    </form>}
    {sortOpen && <form className="sheet-lifecycle" aria-label="Sort selected range" onSubmit={event => { event.preventDefault(); void execute([{kind:'range.sort',range,key_column:sortColumn,descending:sortDescending,header:sortHeader}], 'Selected records sorted'); }}><span>Sort {selectionLabel(selection)}</span><label>Key column<select aria-label="Sort key column" value={sortColumn} onChange={event => setSortColumn(Number(event.target.value))}>{Array.from({length:Math.min(100,range.end_column-range.column+1)},(_,index)=>range.column+index).map(column=><option key={column} value={column}>{address({row:0,column}).replace(/\d+$/, '')}</option>)}</select></label><label><input type="checkbox" checked={sortHeader} onChange={event=>setSortHeader(event.target.checked)}/>First row is a header</label><label>Order<select aria-label="Sort order" value={sortDescending ? 'descending' : 'ascending'} onChange={event=>setSortDescending(event.target.value==='descending')}><option value="ascending">Ascending</option><option value="descending">Descending</option></select></label><button disabled={disabled || sortColumn < range.column || sortColumn > range.end_column} type="submit">Sort records</button><button disabled={busy} type="button" onClick={()=>setSortOpen(false)}>Cancel</button></form>}
    {structureAction && <form className="sheet-lifecycle" aria-label="Row and column changes" onSubmit={event => { event.preventDefault(); applyStructure(); }}><span>{structureAction.endsWith('.delete') ? 'Delete selected whole' : 'Insert before selected'} {structureAction.startsWith('row.') ? `rows ${range.row + 1}–${range.end_row + 1}` : `columns ${address({row:0,column:range.column}).replace(/\d+$/, '')}–${address({row:0,column:range.end_column}).replace(/\d+$/, '')}`}? {structureAction.endsWith('.delete') ? 'Contents in these rows or columns will be removed. You can undo this change.' : 'Cells and references will move together.'}</span><button disabled={disabled} type="submit">{structureAction.endsWith('.delete') ? 'Delete' : 'Insert'}</button><button disabled={busy} type="button" onClick={() => setStructureAction(null)}>Cancel</button></form>}
    {sheetAction && <form className="sheet-lifecycle" aria-label="Worksheet management" onSubmit={event => { event.preventDefault(); void execute([sheetAction === 'delete' ? { kind: 'sheet.delete' } : { kind: sheetAction === 'add' ? 'sheet.add' : 'sheet.rename', name: sheetName }], sheetAction === 'delete' ? 'Worksheet deleted' : sheetAction === 'add' ? 'Worksheet created' : 'Worksheet renamed'); }}>
      {sheetAction === 'delete' ? <span>Delete “{sheet?.name}” and all its contents? Unused ordinary names scoped to or targeting it will also be removed. You can undo this in the current session.</span> : <label>{sheetAction === 'add' ? 'New worksheet name' : 'Rename worksheet'}<input aria-label="Worksheet name" maxLength={31} autoFocus value={sheetName} onChange={event => setSheetName(event.target.value)}/></label>}
      <button disabled={disabled || (sheetAction === 'delete' ? Boolean(deleteSheetReason) : sheetAction==='add' ? Boolean(addSheetReason)||!sheetName.trim() : !sheetName.trim())} type="submit">{sheetAction === 'delete' ? 'Delete worksheet' : sheetAction === 'add' ? 'Create worksheet' : 'Rename worksheet'}</button><button type="button" disabled={busy} onClick={() => setSheetAction(null)}>Cancel</button>
    </form>}
    {!freezeSupported && <div className="sheet-error" role="status">This saved frozen pane exceeds the display limit of 50 rows and 10 columns. Unfreeze or choose a smaller pane to pin it in this editor.</div>}
    {error && <div className="sheet-error" role="alert">{error}</div>}
    <div className="sheet-content"><div className="sheet-grid-scroll" ref={grid} tabIndex={0} role="grid" aria-label="Worksheet cells" aria-rowcount={MAX_ROWS} aria-colcount={MAX_COLUMNS} aria-activedescendant={`sheet-cell-${sheetId}-${address(selection.anchor)}`} onKeyDown={gridKeys}
      onPointerUp={() => { dragging.current = false; }} onPointerLeave={() => { dragging.current = false; }}
      onPaste={event => { if (event.target !== event.currentTarget || draftRef.current !== null || locked.current) return; event.preventDefault(); try { const pasted = pasteOperations(selection.anchor, event.clipboardData.getData('text/plain')); void execute(pasted.operations, 'Pasted cells applied').then(applied => { if (applied) { setSelection(pasted.selection); setLocation(selectionLabel(pasted.selection)); } }); } catch (reason) { setError(String(reason)); } }}
      onCopy={event => { if (event.target !== event.currentTarget || !sheet) return; try { event.clipboardData.setData('text/plain', copySelection(sheet, selection)); event.preventDefault(); setNotice('Selection copied'); } catch (reason) { setError(String(reason)); } }}>
      {!snapshot ? <div className="sheet-loading">{busy ? 'Opening workbook…' : 'The workbook could not be opened.'}</div> : <table className="sheet-grid" style={{ zoom: Math.max(50, Math.min(200, props.viewOptions?.zoom ?? 100)) / 100, width: 48 + columns.reduce((sum, column) => sum + columnPixels(column), 0) }} role="presentation"><colgroup><col style={{ width: 48 }}/>{columns.map(column => <col key={column} style={{ width: columnPixels(column) }}/>)}</colgroup><thead><tr role="row"><th className="sheet-corner" aria-hidden="true"/>{columns.map(column => <th key={column} scope="col" role="columnheader" style={column < frozenColumns ? {...pinnedCellStyle(-1,column),top:0,zIndex:6,backgroundColor:'#f1f5f2'} : undefined}>{address({ row: 0, column }).replace(/\d+$/, '')}</th>)}</tr></thead><tbody>{rows.map(row => <tr key={row} data-sheet-row={row} role="row" aria-rowindex={row + 1} style={{ height: Math.max(24, (sheet?.rows.find(value => value.row === row)?.height_points ?? 20) * 96 / 72) }}><th scope="row" role="rowheader" style={row < frozenRows ? {...pinnedCellStyle(row,-1),left:0,zIndex:6,backgroundColor:'#f1f5f2'} : undefined}>{row + 1}</th>{columns.map(column => {
        const position = { row, column }, merged = sheet?.merged_ranges.find(value => contains(value, position));
        if (merged && (row !== rows.find(value => value >= merged.row) || column !== columns.find(value => value >= merged.column))) return null;
        const source = merged ? { row: merged.row, column: merged.column } : position, key = address(source), cell = cells.get(key), display = cellDisplay(workbook!, cell), calculation = calculatedCells.get(`${sheetId}:${source.row}:${source.column}`), effective = workbook?.styles.find(value => value.id === cell?.style_id)?.effective;
        if (calculation) { display.note = calculation.message ?? 'Calculated locally for this workbook revision.'; if (calculation.status === 'unsupported' || calculation.status === 'circular') display.text = calculation.status === 'circular' ? '#CIRCULAR!' : '#UNSUPPORTED'; }
        const selected = contains(range, position), isActive = address(selection.anchor) === key || Boolean(merged && contains(merged, selection.anchor));
        const css: CSSProperties = { ...pinnedCellStyle(row,column), fontFamily: effective?.font_name, fontSize: effective?.font_size_points ? `${effective.font_size_points}pt` : undefined, fontWeight: effective?.bold ? 700 : undefined, fontStyle: effective?.italic ? 'italic' : undefined, color: effective?.font_color, backgroundColor: effective?.fill_color ?? (row < frozenRows || column < frozenColumns ? '#fff' : undefined), textAlign: effective?.horizontal_alignment === 'general' || !effective?.horizontal_alignment ? cell?.value?.kind === 'number' ? 'right' : 'left' : effective.horizontal_alignment as CSSProperties['textAlign'], verticalAlign: effective?.vertical_alignment === 'middle' ? 'middle' : effective?.vertical_alignment, whiteSpace: effective?.wrap_text ? 'pre-wrap' : 'pre' };
        for (const [edge, side] of Object.entries(effective?.border ?? {})) {
          if (!['top','bottom','left','right'].includes(edge) || !side || typeof side !== 'object' || !('style' in side)) continue;
          const stroke=side as {style:string;color:string}, width=stroke.style==='thick'||stroke.style==='double'?3:stroke.style.startsWith('medium')?2:1;
          (css as Record<string,unknown>)[`border${edge[0]!.toUpperCase()}${edge.slice(1)}`]=`${width}px ${stroke.style==='double'?'double':stroke.style.includes('dott')?'dotted':stroke.style.toLowerCase().includes('dash')?'dashed':'solid'} ${stroke.color}`;
        }
        return <td key={column} id={`sheet-cell-${sheetId}-${key}`} role="gridcell" aria-colindex={column + 1} aria-selected={selected} aria-readonly={!sheet?.editable || Boolean(merged) || cell?.editable === false} className={`${selected ? 'is-selected' : ''} ${isActive ? 'is-active' : ''} ${cell?.formula ? 'has-formula' : ''}`} style={css} rowSpan={merged ? rows.filter(value => value >= merged.row && value <= merged.end_row).length : undefined} colSpan={merged ? columns.filter(value => value >= merged.column && value <= merged.end_column).length : undefined} title={`${key}${merged ? ` · merged ${merged.ref}` : ''}${display.note ? ` · ${display.note}` : ''}`} onPointerDown={event => { if (event.button !== 0) return; event.preventDefault(); dragging.current = true; select(source, event.shiftKey); grid.current?.focus(); }} onPointerEnter={event => { if (dragging.current && event.buttons === 1) select(position, true); }} onDoubleClick={() => { if (address(selection.anchor) === key) startEdit(); }}>
          {isActive && inline && draft !== null ? <input ref={inlineInput} aria-label={`Edit ${key}`} className="sheet-inline-input" maxLength={32767} onCompositionStart={() => { composing.current = true; }} onCompositionEnd={() => { composing.current = false; }} value={draft} disabled={busy} onPointerDown={event => event.stopPropagation()} onChange={event => updateDraft(event.target.value)} onKeyDown={editKeys}/> : <span className="sheet-cell-content">{display.text}</span>}
        </td>;
      })}</tr>)}</tbody></table>}
    </div>
    {chartsOpen && snapshot && <SpreadsheetCharts charts={snapshot.charts} error={snapshot.chartError} sheetId={sheetId} range={range} disabled={disabled} onExecute={execute} onClose={() => setChartsOpen(false)}/>}
    </div>
    <div className="sheet-bottom"><div className="sheet-tab-tools"><button aria-label="Add worksheet" title={addSheetReason} disabled={disabled || Boolean(addSheetReason)} onClick={() => { let number = 1; while (workbook?.sheets.some(value => value.name.toLowerCase() === `sheet${number}`)) number++; setSheetName(`Sheet${number}`); setSheetAction('add'); }}>+</button><button disabled={disabled} onClick={() => { setSheetName(sheet?.name ?? ''); setSheetAction('rename'); }}>Rename</button><button title={deleteSheetReason} disabled={disabled || Boolean(deleteSheetReason)} onClick={() => setSheetAction('delete')}>Delete</button></div><nav aria-label="Worksheets">{workbook?.sheets.filter(value => value.state === 'visible').map(value => <button key={value.id} aria-current={value.id === sheetId ? 'page' : undefined} disabled={busy || draft !== null} onClick={() => { setSheetId(value.id); setSelection(initialSelection); setLocation('A1'); setOrigin({ row: 0, column: 0 }); }}>{value.name}{!value.editable ? ' · read-only' : ''}</button>)}</nav><div className="sheet-window-controls"><button aria-label="Previous rows" disabled={busy || draft !== null || origin.row === 0} onClick={() => select({ row: Math.max(0, origin.row - 60), column: origin.column })}>↑</button><button aria-label="Previous columns" disabled={busy || draft !== null || origin.column === 0} onClick={() => select({ row: origin.row, column: Math.max(0, origin.column - 20) })}>←</button><span>Rows {origin.row + 1}–{Math.min(MAX_ROWS, origin.row + 60)}</span><button aria-label="Next rows" disabled={busy || draft !== null || origin.row + 60 >= MAX_ROWS} onClick={() => select({ row: origin.row + 60, column: origin.column })}>↓</button><button aria-label="Next columns" disabled={busy || draft !== null || origin.column + 20 >= MAX_COLUMNS} onClick={() => select({ row: origin.row, column: origin.column + 20 })}>→</button></div></div>
    <div className="sheet-status"><span>{busy ? 'Applying native change…' : activeStatus}</span><span role="status">{notice || 'Enter / F2 to edit · Shift + arrows to select · paste a range'}</span><span>{snapshot?.calculation ? `Local calculation · ${snapshot.calculation.cells.filter(cell => cell.status === 'unsupported' || cell.status === 'circular').length} unresolved` : 'Formula caches: stored, unverified'}</span></div>
  </section>;
}
