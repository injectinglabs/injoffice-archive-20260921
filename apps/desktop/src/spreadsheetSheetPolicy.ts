import type { NativeWorkbookV2 } from '@injoffice/sheets/browser';

/** Built-in Excel names are prefixed with `_`; they are not user-editable. */
function editableDefinedName(name: string): boolean {
  return !String(name).startsWith('_');
}
/** Predictable native refusals, shown before presenting a destructive action. */
export function sheetLifecycleReason(workbook: NativeWorkbookV2 | undefined, sheetId: string, action: 'add' | 'delete'): string | undefined {
  const target = workbook?.sheets.find(sheet => sheet.id === sheetId);
  if (!workbook || !target?.editable) return 'This worksheet is read-only.';
  const blocked = new Set(['drawings', 'tables', 'external-links', 'extensions', 'conditional-formatting', 'data-validation', 'hyperlinks', 'protection']);
  if (workbook.unsupported.some(item => blocked.has(item.capability))) return 'Worksheet management requires a workbook without unqualified drawing, validation, link or worksheet features.';
  if(workbook.sheets.some(sheet=>sheet.auto_filter))return 'Clear worksheet filters before changing worksheet topology.';
  if (action === 'add') {
    if (workbook.sheets.length >= 1024 || workbook.sheets.some(sheet => Number(sheet.id) >= 4294967295)) return 'The workbook has reached its worksheet identity limit.';
    if (workbook.sheets.some(sheet => !sheet.editable && sheet.cells.some(cell => cell.formula)) || workbook.sheets.reduce((n,sheet)=>n+sheet.cells.filter(cell=>cell.formula).length,0)>10000) return 'Formula caches cannot be safely invalidated when appending this worksheet.';
    return;
  }
  if (workbook.sheets.length < 2 || target.state === 'visible' && workbook.sheets.filter(sheet => sheet.state === 'visible').length < 2) return 'Keep at least one visible worksheet.';
  if (workbook.sheets.some(sheet => sheet.cells.some(cell => cell.formula))) return 'Delete is unavailable while the workbook contains formulas; reference-aware deletion is not yet supported.';
  for (const name of workbook.defined_names ?? []) {
    if (!name.topology_safe || !name.target_sheet_id) return 'Delete is unavailable while advanced defined names have unqualified dependencies.';
    if ((name.target_sheet_id === sheetId || name.scope_sheet_id === sheetId) && (name.hidden || !editableDefinedName(name.name))) return 'An affected hidden or built-in name must be preserved.';
  }
}
