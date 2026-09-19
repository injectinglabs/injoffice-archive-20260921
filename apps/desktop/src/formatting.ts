import type { DocumentTextRange } from './document-range';
import type { NativeWorkbookV2, WorkbookMutationBatchV1, StyleDelta } from '@injoffice/sheets/browser';
import type { NativeDocxDocumentV1, NativeDocxParagraphV1 } from '../../../packages/docs/src/nativeContract';
import type { NativeDocxOfficeMutationEnvelopeV1, NativeDocxRunPropertyPatchV1, NativeDocxParagraphPropertyPatchV1 } from '../../../packages/docs/src/nativeTransactionAdapterV1';
import type { NativePptxDeck } from '@injoffice/pptx-native';
import type { PptxNativeMutationRequestV1 } from '@injoffice/pptx-wasm';
import { editableTargets, targetKey } from '../../playground/src/nativeRoundTrip';
import { editableDocxRuns } from '../../playground/src/docxRoundTrip';
import { editablePptxTextTargets, pptxTargetKey } from '../../playground/src/pptxRoundTrip';
import { paragraphAppearance, runAppearance } from './document-style';
import type { FormattingPatch, FormattingValues } from './FormattingToolbar';

export type FormattingPreview = { kind: 'docx'; document: NativeDocxDocumentV1 } | { kind: 'xlsx'; workbook: NativeWorkbookV2 } | { kind: 'pptx'; deck: NativePptxDeck } | { kind: 'pdf' };
export function docxSelection(document: NativeDocxDocumentV1, key: string) {
  const target = editableDocxRuns(document).find(value => value.key === key);
  if (!target) return;
  const paragraphs: NativeDocxParagraphV1[] = [];
  for (const story of [document.body, ...document.headers, ...document.footers, ...document.notes, ...document.comment_stories]) {
    for (const block of story.blocks) {
      if (block.paragraph) paragraphs.push(block.paragraph);
      else for (const row of block.table?.rows ?? []) for (const cell of row.cells) paragraphs.push(...cell.paragraphs);
    }
  }
  const paragraph = paragraphs.find(value => value.id === target.paragraphId);
  const run = paragraph?.runs.find(value => value.id === target.runId);
  return paragraph && run ? { target, paragraph, run } : undefined;
}
export function formattingValues(preview: FormattingPreview, key: string): FormattingValues | undefined {
  if (preview.kind === 'docx') {
    const selected = docxSelection(preview.document, key);
    if (!selected || !selected.paragraph.edit_policy.allowed_operations.includes('properties.patch')) return;
    const properties = runAppearance(preview.document, selected.paragraph, selected.run);
    const ownerPath = selected.run.anchor.path.slice(0, selected.run.anchor.path.lastIndexOf('/'));
    const siblings = selected.paragraph.runs.filter(run => run.anchor.path.slice(0, run.anchor.path.lastIndexOf('/')) === ownerPath);
    return { font: properties.font_family, size: properties.font_size_half_points ? properties.font_size_half_points / 2 : undefined, bold: properties.bold, italic: properties.italic, underline: properties.underline === undefined ? undefined : properties.underline !== 'none', color: properties.color && /^[A-Fa-f0-9]{6}$/.test(properties.color) ? `#${properties.color.toUpperCase()}` : undefined, alignment: paragraphAppearance(preview.document, selected.paragraph).paragraph.alignment, characterEditable: siblings.length === 1 && selected.run.kind === 'text' };
  }
  if (preview.kind === 'xlsx') {
    const cell = editableTargets(preview.workbook).find(value => targetKey(value) === key);
    const properties = preview.workbook.styles.find(value => value.id === cell?.style_id)?.effective;
    if (!cell || !properties || properties.projection !== 'full') return;
    return { font: properties.font_name || 'Arial', size: properties.font_size_points || 11, bold: Boolean(properties.bold), italic: Boolean(properties.italic), color: properties.font_color || '#20242B', alignment: properties.horizontal_alignment || 'general', numberFormat: properties.number_format || 'General' };
  }
  if (preview.kind === 'pptx') {
    const target = editablePptxTextTargets(preview.deck).find(value => pptxTargetKey(value) === key);
    const paragraph = target?.paragraphs[0], run = paragraph?.runs[0];
    if (!run) return;
    return { font: run.fontFamily!, size: run.fontSizeHundredthPt! / 100, bold: Boolean(run.bold), italic: Boolean(run.italic), color: `#${run.color}`, alignment: paragraph!.align! };
  }
}
/** Aggregate only the selected source text, including an uncommitted active-run draft. */
export function documentRangeFormattingValues(document:NativeDocxDocumentV1,key:string,range:DocumentTextRange,draft?:string):FormattingValues|undefined {
 const selected=docxSelection(document,key),base=formattingValues({kind:'docx',document},key)
 if(!selected||!base||range.paragraph_id!==selected.paragraph.id)return base
 let offset=0
 const runs=selected.paragraph.runs.filter(run=>{const start=offset;offset+=(run.id===selected.run.id&&draft!==undefined?draft:run.text??'').length;return start<range.end_utf16&&offset>range.start_utf16})
 const values=runs.map(run=>{const properties=runAppearance(document,selected.paragraph,run);return {font:properties.font_family,size:properties.font_size_half_points===undefined?undefined:properties.font_size_half_points/2,bold:properties.bold??false,italic:properties.italic??false,underline:properties.underline!==undefined&&properties.underline!=='none',color:properties.color&&/^[a-f0-9]{6}$/i.test(properties.color)?`#${properties.color.toUpperCase()}`:undefined}})
 const result:FormattingValues={...base,characterEditable:!!selected.paragraph.can_format_range&&!!values.length&&!range.unsupported&&!selected.paragraph.runs.some(run=>runAppearance(document,selected.paragraph,run).hidden)}
 for(const field of ['font','size','bold','italic','underline','color'] as const){const value=values[0]?.[field];Object.assign(result,{[field]:values.every(item=>item[field]===value)?value:undefined})}
 return result
}
export function workbookFormatting(workbook: NativeWorkbookV2, key: string, patch: FormattingPatch, id: string): WorkbookMutationBatchV1 {
  const cell = editableTargets(workbook).find(value => targetKey(value) === key);
  if (!cell) throw new Error('Select an editable cell first.');
  const style: StyleDelta = {};
  if (patch.font !== undefined) style.font_name = patch.font;
  if (patch.size !== undefined) style.font_size_points = patch.size;
  if (patch.bold !== undefined) style.bold = patch.bold;
  if (patch.italic !== undefined) style.italic = patch.italic;
  if (patch.color !== undefined) style.font_color = patch.color;
  if (patch.numberFormat !== undefined) style.number_format = patch.numberFormat;
  if (patch.alignment !== undefined) style.horizontal_alignment = patch.alignment as StyleDelta['horizontal_alignment'];
  return { protocol: 'injoffice.xlsx.mutations', version: 1, batch_id: id, expected_revision: workbook.source.package_sha256, operations: [{ operation_id: id, sheet_id: cell.sheetId, kind: 'style.patch', range: { row: cell.row, column: cell.column, end_row: cell.row, end_column: cell.column }, style }] };
}
export function presentationFormatting(deck: NativePptxDeck, key: string, patch: FormattingPatch, id: string): PptxNativeMutationRequestV1 {
  const target = editablePptxTextTargets(deck).find(value => pptxTargetKey(value) === key);
  if (!target || !deck.sourceRevision) throw new Error('Select editable text first.');
  const paragraphs = target.paragraphs.map((paragraph, p) => ({ ...paragraph, ...(p === 0 && patch.alignment !== undefined ? { align: patch.alignment as 'left' | 'center' | 'right' } : {}), runs: paragraph.runs.map((run, r) => p || r ? run : { ...run,
    ...(patch.font !== undefined ? { fontFamily: patch.font } : {}), ...(patch.size !== undefined ? { fontSizeHundredthPt: Math.round(patch.size * 100) } : {}), ...(patch.bold !== undefined ? { bold: patch.bold } : {}), ...(patch.italic !== undefined ? { italic: patch.italic } : {}), ...(patch.color !== undefined ? { color: patch.color.replace(/^#/, '').toUpperCase() } : {}),
  }) }));
  return { expectedSourceRevision: deck.sourceRevision, operations: [{ operationId: id, kind: 'text.replace', elementId: target.elementId, expectedFingerprintSha256: target.expectedFingerprintSha256, paragraphs }] };
}
export function documentFormatting(document: NativeDocxDocumentV1, key: string, patch: FormattingPatch, id: string, range?: DocumentTextRange): NativeDocxOfficeMutationEnvelopeV1 {
  const selected = docxSelection(document, key);
  if (!selected) throw new Error('Select editable document text first.');
  const isParagraph = patch.alignment !== undefined;
  if (isParagraph && Object.keys(patch).length !== 1) throw new Error('Apply paragraph alignment separately from text formatting.');
  const paragraphRange=!isParagraph&&!!range?.paragraph_id;
  if(paragraphRange&&(range!.paragraph_id!==selected.paragraph.id||!selected.paragraph.can_format_range||range!.unsupported||selected.paragraph.runs.some(run=>runAppearance(document,selected.paragraph,run).hidden)))throw new Error('This paragraph selection cannot be formatted safely.');
  if(range && !isParagraph && !paragraphRange && !selected.run.can_format_range) throw new Error('Selected-text formatting is not supported in this run.');
  const properties: NativeDocxRunPropertyPatchV1 = {};
  if (patch.font !== undefined) properties.font_family = patch.font;
  if (patch.size !== undefined) properties.font_size_half_points = Math.round(patch.size * 2);
  if (patch.bold !== undefined) properties.bold = patch.bold;
  if (patch.italic !== undefined) properties.italic = patch.italic;
  if (patch.underline !== undefined) properties.underline = patch.underline ? 'single' : 'none';
  if (patch.color !== undefined) properties.color = patch.color.replace(/^#/, '').toUpperCase();
  const target = isParagraph || paragraphRange ? selected.paragraph : selected.run;
  const anchor = { target_id: target.id, expected_xml_sha256: target.anchor.xml_sha256 };
  const mutation = isParagraph ? { ...anchor, target_kind: 'paragraph' as const, properties: { alignment: patch.alignment as NativeDocxParagraphPropertyPatchV1['alignment'] } } : paragraphRange ? {...anchor,target_kind:'paragraph' as const,properties,range:{start_utf16:range!.start_utf16,end_utf16:range!.end_utf16}} : { ...anchor, target_kind: 'run' as const, properties, ...(range ? {range:{start_utf16:range.start_utf16,end_utf16:range.end_utf16}} : {}) };
  return { protocol: 'injoffice.office.mutations' as const, version: 1 as const, format: 'docx' as const, mutation_id: id, expected_revision: document.source.package_sha256,
    payload: { mutations: [mutation] } };
}

export type ParagraphOperation = 'block.insert_after' | 'block.delete';
export function paragraphOperations(document: NativeDocxDocumentV1, key: string): ParagraphOperation[] {
  const selected = docxSelection(document, key);
  if (!selected || selected.paragraph.edit_policy.mode !== 'read-write') return [];
  return (['block.insert_after', 'block.delete'] as const).filter(operation => selected.paragraph.edit_policy.allowed_operations.includes(operation));
}
export function documentStructure(document: NativeDocxDocumentV1, key: string, operation: ParagraphOperation, id: string): NativeDocxOfficeMutationEnvelopeV1 {
  const selected = docxSelection(document, key);
  if (!selected || !paragraphOperations(document, key).includes(operation)) throw new Error('This paragraph does not support the selected structure command.');
  const target = selected.paragraph;
  const anchor = { target_kind: 'paragraph' as const, target_id: target.id, expected_xml_sha256: target.anchor.xml_sha256 };
  const mutation = operation === 'block.insert_after' ? { ...anchor, operation, text: '' } : { ...anchor, operation };
  return { protocol: 'injoffice.office.mutations' as const, version: 1 as const, format: 'docx' as const, mutation_id: id, expected_revision: document.source.package_sha256,
    payload: { mutations: [mutation] } };
}

export function documentParagraphFormatting(document: NativeDocxDocumentV1, key: string, properties: NativeDocxParagraphPropertyPatchV1, id: string): NativeDocxOfficeMutationEnvelopeV1 {
  const selected = docxSelection(document,key);
  if (!selected?.paragraph.edit_policy.allowed_operations.includes('properties.patch')) throw new Error('Select a paragraph with supported formatting.');
  return {protocol:'injoffice.office.mutations',version:1,format:'docx',mutation_id:id,expected_revision:document.source.package_sha256,payload:{mutations:[{target_kind:'paragraph',target_id:selected.paragraph.id,expected_xml_sha256:selected.paragraph.anchor.xml_sha256,properties}]}};
}

export function documentTableSelection(document: NativeDocxDocumentV1, key: string) {
  const selected=docxSelection(document,key)
  return selected ? document.body.blocks.find(block=>block.table?.rows.some(row=>row.cells.some(cell=>cell.paragraphs.some(paragraph=>paragraph.id===selected.paragraph.id))))?.table : undefined
}
