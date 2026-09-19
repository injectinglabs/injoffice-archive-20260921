import {imageDimensions} from './raster-image'
export {imageDimensions} from './raster-image'
import type { NativeElement, NativePptxDeck, NativeTransform } from '@injoffice/pptx-native';
import type { PptxNativeExactAutoShapeV1, PptxNativeExactParagraphV1, PptxNativeMutationRequestV1 } from '@injoffice/pptx-wasm';
import { editablePptxTextTargets, editablePptxShapeTargets } from '../../playground/src/pptxRoundTrip';

export const EMU_PER_INCH = 914400;
export const EMU_PER_PIXEL = 9525;
export function elementKey(element: NativeElement): string { return element.source ? `${element.source.partName}\0${element.source.objectId}` : element.id; }
export function textTarget(deck: NativePptxDeck, key: string) { return editablePptxTextTargets(deck).find(target => `${target.sourcePartName}\0${target.sourceObjectId}` === key); }
export function shapeTarget(deck: NativePptxDeck, key: string) { return editablePptxShapeTargets(deck).find(target => `${target.sourcePartName}\0${target.sourceObjectId}` === key); }
function revision(deck: NativePptxDeck) { if (!deck.sourceRevision) throw new Error('Open a source presentation before editing.'); return deck.sourceRevision; }
export function textCommand(deck: NativePptxDeck, key: string, paragraphs: readonly PptxNativeExactParagraphV1[], operationId: string): PptxNativeMutationRequestV1 {
  const target = textTarget(deck, key); if (!target) throw new Error('This text is preserved but is not editable.');
  if (JSON.stringify(target.paragraphs) === JSON.stringify(paragraphs)) throw new Error('There are no text changes to apply.');
  return { expectedSourceRevision: revision(deck), operations: [{ operationId, kind: 'text.replace', elementId: target.elementId, expectedFingerprintSha256: target.expectedFingerprintSha256, paragraphs: structuredClone(paragraphs) }] };
}
export function shapeCommand(deck: NativePptxDeck, key: string, autoShape: PptxNativeExactAutoShapeV1, operationId: string): PptxNativeMutationRequestV1 {
  const target = shapeTarget(deck, key); if (!target) throw new Error('This shape is preserved but is not editable.');
  for (const [name, value] of Object.entries(autoShape.transform)) if (!Number.isSafeInteger(value) || ((name === 'cx' || name === 'cy') && value <= 0)) throw new Error('Use positive dimensions and valid positions.');
  if (JSON.stringify(target.autoShape) === JSON.stringify(autoShape)) throw new Error('There are no shape changes to apply.');
  return { expectedSourceRevision: revision(deck), operations: [{ operationId, kind: 'autoshape.update', elementId: target.elementId, expectedFingerprintSha256: target.expectedFingerprintSha256, autoShape: structuredClone(autoShape) }] };
}
export function structureCommand(deck: NativePptxDeck, index: number, action: 'duplicate' | 'delete' | 'previous' | 'next', operationId: string): PptxNativeMutationRequestV1 {
  const slide = deck.slides[index]; if (!slide) throw new Error('Select a slide first.');
  const expectedSourceRevision = revision(deck);
  if (action === 'delete' || action === 'duplicate') {
    if (action === 'delete' && deck.slides.length <= 1) throw new Error('The last slide cannot be deleted.');
    return { expectedSourceRevision, operations: [{ operationId, kind: action === 'delete' ? 'slide.delete' : 'slide.duplicate', slideId: slide.id }] };
  }
  const destination = index + (action === 'previous' ? -1 : 1);
  if (destination < 0 || destination >= deck.slides.length) throw new Error('This slide is already at the end of the presentation.');
  const slideIds = deck.slides.map(value => value.id); [slideIds[index], slideIds[destination]] = [slideIds[destination]!, slideIds[index]!];
  return { expectedSourceRevision, operations: [{ operationId, kind: 'slide.reorder', slideIds }] };
}
export interface PositionedElement { element: NativeElement; rect: NativeTransform; scaleX: number; scaleY: number; grouped: boolean }
/** Project the native group's child coordinate space for preview only. Saving uses original local coordinates. */
export function positionElements(elements: readonly NativeElement[], parent = { x: 0, y: 0, sx: 1, sy: 1 }, grouped = false): PositionedElement[] {
  return elements.flatMap(element => {
    const t = element.transform;
    const rect = { x: parent.x + t.x * parent.sx, y: parent.y + t.y * parent.sy, cx: t.cx * parent.sx, cy: t.cy * parent.sy };
    if (element.kind === 'group' && element.compatibility.status !== 'refused') {
      const child = element.childTransform ?? { x: 0, y: 0, cx: t.cx, cy: t.cy };
      if (!child.cx || !child.cy) return [{ element, rect, scaleX: parent.sx, scaleY: parent.sy, grouped }];
      const sx = parent.sx * t.cx / child.cx, sy = parent.sy * t.cy / child.cy;
      return positionElements(element.children, { x: rect.x - child.x * sx, y: rect.y - child.y * sy, sx, sy }, true);
    }
    return [{ element, rect, scaleX: parent.sx, scaleY: parent.sy, grouped }];
  });
}

export type PresentationDraft = { kind: 'rotation'; rotation60000: number } | { kind: 'table'; row: number; column: number; paragraphs: PptxNativeExactParagraphV1[]; fill?: string } | { kind: 'geometry'; transform: NativeTransform } | { kind: 'text'; paragraphs: PptxNativeExactParagraphV1[] } | { kind: 'shape'; shape: PptxNativeExactAutoShapeV1 };
export interface PresentationRecoveryDraft { version: 1; format: 'pptx'; revision: string; key: string; draft: PresentationDraft }
const MAX_RECOVERY_JSON = 256 * 1024;
/** Check shape/types before displaying recovered JSON. Native Apply still validates all mutation semantics. */
function conform(value: unknown, template: unknown): boolean {
  if (typeof template === 'number') return value === null || (typeof value === 'number' && Number.isFinite(value));
  if (typeof template === 'string') return typeof value === 'string';
  if (typeof template === 'boolean') return typeof value === 'boolean';
  if (Array.isArray(template)) return Array.isArray(value) && value.length === template.length && value.every((item, i) => conform(item, template[i]));
  if (template && typeof template === 'object') {
    if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
    const record = value as Record<string, unknown>; const sample = template as Record<string, unknown>;
    return Object.keys(record).every(key => Object.hasOwn(sample, key)) && Object.keys(sample).every(key => conform(record[key], sample[key]));
  }
  return value === template;
}
export function recoveryDraft(deck: NativePptxDeck, key: string, draft: PresentationDraft): PresentationRecoveryDraft {
  const value: PresentationRecoveryDraft = { version: 1, format: 'pptx', revision: revision(deck), key, draft };
  const json = JSON.stringify(value);
  if (json.length > MAX_RECOVERY_JSON) throw new Error('This pending edit is too large for recovery. Apply it to include it in the document checkpoint.');
  return JSON.parse(json);
}
export function restorePresentationDraft(deck: NativePptxDeck, input: unknown): PresentationRecoveryDraft | undefined {
  if (input === undefined || input === null) return undefined;
  const json = JSON.stringify(input);
  if (!json || json.length > MAX_RECOVERY_JSON) throw new Error('The recovered presentation edit exceeds the supported size.');
  const value = JSON.parse(json) as PresentationRecoveryDraft;
  if (value?.version !== 1 || value.format !== 'pptx' || value.revision !== deck.sourceRevision || typeof value.key !== 'string' || !value.draft) throw new Error('The recovered edit does not match this presentation revision.');
  if (value.draft.kind === 'rotation') {
    const angle = value.draft.rotation60000; if (!rotationTarget(deck, value.key) || !Number.isSafeInteger(angle) || angle < 0 || angle >= 21600000) throw new Error('The recovered rotation is invalid.');
  } else if (value.draft.kind === 'text') {
    const target = textTarget(deck, value.key);
    if (!target || !conform(value.draft.paragraphs, target.paragraphs)) throw new Error('The recovered text target or formatting is invalid.');
    value.draft.paragraphs = value.draft.paragraphs.map(p => ({ ...p, runs: p.runs.map(r => ({ ...r, fontSizeHundredthPt: r.fontSizeHundredthPt ?? NaN })) }));
  } else if (value.draft.kind === 'table') {
    const pending = value.draft; const table = tableTarget(deck, value.key);
    const cell = table?.element.table.rows[pending.row]?.[pending.column];
    const sample = { align: 'left', level: 0, bullet: false, runs: [{ text: '', bold: false, italic: false, fontFamily: 'Arial', fontSizeHundredthPt: 1800, color: '20242B' }] };
    if (!cell || !Number.isSafeInteger(pending.row) || !Number.isSafeInteger(pending.column) || !Array.isArray(pending.paragraphs) || !pending.paragraphs.length || pending.paragraphs.length > 1000 || !pending.paragraphs.every(p => p && Array.isArray(p.runs) && p.runs.length > 0 && conform({ ...p, runs: [p.runs[0]] }, sample) && p.runs.every(r => conform(r, sample.runs[0]))) || (pending.fill !== undefined && typeof pending.fill !== 'string')) throw new Error('The recovered table cell edit is invalid.');
  } else if (value.draft.kind === 'geometry') {
    const target = deck.slides.flatMap(slide => positionElements(slide.elements)).find(item => elementKey(item.element) === value.key)?.element;
    if (!target || !['text', 'picture', 'table'].includes(target.kind) || target.compatibility.status !== 'editable' || !conform(value.draft.transform, target.transform)) throw new Error('The recovered geometry target is invalid.');
    for (const key of ['x', 'y', 'cx', 'cy'] as const) value.draft.transform[key] ??= NaN;
  } else if (value.draft.kind === 'shape') {
    const target = shapeTarget(deck, value.key); const shape = value.draft.shape;
    if (!target || !shape || !conform(shape.transform, target.autoShape.transform) || !['rect', 'ellipse', 'triangle', 'diamond'].includes(shape.preset) || (shape.fill !== undefined && typeof shape.fill !== 'string')) throw new Error('The recovered shape target or geometry is invalid.');
    if (Object.keys(shape).some(key => !['transform', 'preset', 'fill', 'stroke'].includes(key))) throw new Error('The recovered shape has unsupported fields.');
    if (shape.stroke) {
      const { miterLimit, ...stroke } = shape.stroke;
      if (!conform(stroke, { color: '', widthEmu: 0, cap: '', join: '', dash: '' }) || (miterLimit !== undefined && typeof miterLimit !== 'number')) throw new Error('The recovered outline is invalid.');
      shape.stroke.widthEmu ??= NaN;
    }
    for (const key of ['x', 'y', 'cx', 'cy'] as const) shape.transform[key] ??= NaN;
  } else throw new Error('The recovered presentation edit kind is unsupported.');
  return value;
}

export function insertCommand(deck: NativePptxDeck, index: number, kind: 'slide' | 'text' | 'shape', operationId: string): PptxNativeMutationRequestV1 {
  const slideId = deck.slides[index]?.id; if (!slideId) throw new Error('Select a slide before inserting.');
  const base = { operationId, slideId }; const expectedSourceRevision = revision(deck);
  const transform = { x: Math.round(deck.size.cx * .15), y: Math.round(deck.size.cy * .2), cx: Math.max(1, Math.round(deck.size.cx * .4)), cy: Math.max(1, Math.round(deck.size.cy * .2)) };
  if (kind === 'slide') return { expectedSourceRevision, operations: [{ ...base, kind: 'slide.insert' }] };
  if (kind === 'shape') return { expectedSourceRevision, operations: [{ ...base, kind: 'autoshape.insert', autoShape: { transform, preset: 'rect', fill: 'DCE8F7', stroke: { color: '2459AD', widthEmu: 12700, cap: 'flat', join: 'round', dash: 'solid' } } }] };
  return { expectedSourceRevision, operations: [{ ...base, kind: 'text.insert', transform, paragraphs: [{ align: 'left', level: 0, bullet: false, runs: [{ text: '', fontFamily: 'Arial', fontSizeHundredthPt: 2400, bold: false, italic: false, color: '20242B' }] }] }] };
}
export function deleteElementCommand(deck: NativePptxDeck, key: string, operationId: string): PptxNativeMutationRequestV1 {
  const element = deck.slides.flatMap(slide => slide.elements).find(value => elementKey(value) === key);
  if (!element?.source || element.compatibility.status !== 'editable' || (element.kind !== 'text' && element.kind !== 'shape' && element.kind !== 'picture' && element.kind !== 'table')) throw new Error('Only exact top-level text, shapes, pictures or tables can be deleted.');
  return { expectedSourceRevision: revision(deck), operations: [{ operationId, kind: 'element.delete', elementId: element.id, expectedFingerprintSha256: element.source.fingerprintSha256 }] };
}

export function geometryTarget(deck: NativePptxDeck, key: string) {
  return deck.slides.flatMap((slide, slideIndex) => positionElements(slide.elements).map(item => ({ ...item, slideIndex }))).find(item => elementKey(item.element) === key && item.element.source && item.element.compatibility.status === 'editable' && (item.element.kind === 'text' || item.element.kind === 'picture' || item.element.kind === 'table'));
}
export function transformCommand(deck: NativePptxDeck, key: string, transform: NativeTransform, operationId: string): PptxNativeMutationRequestV1 {
  const target = geometryTarget(deck, key)?.element;
  if (!target?.source) throw new Error('This object cannot be repositioned.');
  for (const [key, value] of Object.entries(transform)) if (!Number.isSafeInteger(value) || ((key === 'cx' || key === 'cy') && value <= 0)) throw new Error('Use valid positions and positive dimensions.');
  return { expectedSourceRevision: revision(deck), operations: [{ operationId, kind: target.kind === 'picture' ? 'picture.transform' : target.kind === 'table' ? 'table.transform' : 'text.transform', elementId: target.id, expectedFingerprintSha256: target.source.fingerprintSha256, transform: { ...transform } }] };
}

export function pointerTransform(start: NativeTransform, dx: number, dy: number, pixelsPerEmuX: number, pixelsPerEmuY: number, resize: boolean): NativeTransform {
  if (![dx, dy, pixelsPerEmuX, pixelsPerEmuY].every(Number.isFinite) || pixelsPerEmuX <= 0 || pixelsPerEmuY <= 0) return { ...start };
  const x = Math.round(dx / pixelsPerEmuX), y = Math.round(dy / pixelsPerEmuY);
  return resize ? { ...start, cx: Math.max(1, start.cx + x), cy: Math.max(1, start.cy + y) } : { ...start, x: start.x + x, y: start.y + y };
}

export function pictureInsertCommand(deck: NativePptxDeck, index: number, bytes: Uint8Array, operationId: string): PptxNativeMutationRequestV1 {
  const dimensions = imageDimensions(bytes); const slideId = deck.slides[index]?.id; if (!slideId) throw new Error('Select a slide first.');
  const scale = Math.min(deck.size.cx * .7 / dimensions.width, deck.size.cy * .7 / dimensions.height);
  const cx = Math.max(1, Math.round(dimensions.width * scale)), cy = Math.max(1, Math.round(dimensions.height * scale));
  let binary = ''; for (let i = 0; i < bytes.length; i += 32768) binary += String.fromCharCode(...bytes.subarray(i, i + 32768));
  return { expectedSourceRevision: revision(deck), operations: [{ operationId, kind: 'picture.insert', slideId, transform: { x: Math.round((deck.size.cx - cx) / 2), y: Math.round((deck.size.cy - cy) / 2), cx, cy }, image: { contentType: dimensions.contentType, dataBase64: btoa(binary) } }] };
}

export function tableTarget(deck: NativePptxDeck, key: string) {
  const target = geometryTarget(deck, key); return target?.element.kind === 'table' ? { ...target, element: target.element } : undefined;
}
export function tableInsertCommand(deck: NativePptxDeck, index: number, rows: number, columns: number, operationId: string): PptxNativeMutationRequestV1 {
  if (!Number.isSafeInteger(rows) || !Number.isSafeInteger(columns) || rows < 1 || columns < 1 || rows > 100 || columns > 100 || rows * columns > 1000) throw new Error('Choose 1–100 rows and columns, up to 1000 cells.');
  const slideId = deck.slides[index]?.id; if (!slideId) throw new Error('Select a slide first.');
  return { expectedSourceRevision: revision(deck), operations: [{ operationId, kind: 'table.insert', slideId, rows, columns, transform: { x: Math.round(deck.size.cx * .15), y: Math.round(deck.size.cy * .2), cx: Math.max(columns, Math.round(deck.size.cx * .7)), cy: Math.max(rows, Math.round(deck.size.cy * .55)) } }] };
}
export function tableCellCommand(deck: NativePptxDeck, key: string, draft: Extract<PresentationDraft, { kind: 'table' }>, operationId: string): PptxNativeMutationRequestV1 {
  const target = tableTarget(deck, key)?.element;
  if (!target?.source || !target.table.rows[draft.row]?.[draft.column]) throw new Error('This table cell cannot be edited.');
  return { expectedSourceRevision: revision(deck), operations: [{ operationId, kind: 'table.cell.replace', elementId: target.id, expectedFingerprintSha256: target.source.fingerprintSha256, row: draft.row, column: draft.column, paragraphs: structuredClone(draft.paragraphs), ...(draft.fill === undefined ? {} : { fill: draft.fill }) }] };
}
export function tableTextParagraphs(original: readonly PptxNativeExactParagraphV1[], text: string): PptxNativeExactParagraphV1[] {
  const sample = original[0] ?? { align: 'left', level: 0, bullet: false, runs: [] };
  const run = sample.runs[0] ?? { text: '', bold: false, italic: false, fontFamily: 'Arial', fontSizeHundredthPt: 1800, color: '20242B' };
  return text.split(/\r?\n/).map(line => ({ ...sample, runs: [{ ...run, text: line }] }));
}

/** Replace one text segment; preserve all other runs and paragraph metadata.
 * Newlines split this segment into paragraphs with its paragraph formatting. */
export function tableSegmentParagraphs(original: readonly PptxNativeExactParagraphV1[], paragraphIndex: number, runIndex: number, text: string): PptxNativeExactParagraphV1[] {
  const paragraph = original[paragraphIndex]; const run = paragraph?.runs[runIndex];
  if (!run) throw new Error('Select a table text segment.');
  const lines = text.replace(/\r\n?/g, '\n').split('\n');
  const replacement = lines.map((line, index) => ({ ...structuredClone(paragraph), runs: [
    ...(index === 0 ? structuredClone(paragraph.runs.slice(0, runIndex)) : []),
    { ...structuredClone(run), text: line },
    ...(index === lines.length - 1 ? structuredClone(paragraph.runs.slice(runIndex + 1)) : []),
  ] }));
  return [...structuredClone(original.slice(0, paragraphIndex)), ...replacement, ...structuredClone(original.slice(paragraphIndex + 1))];
}

export function tableTopologyCommand(deck: NativePptxDeck, key: string, axis: 'row' | 'column', action: 'insert' | 'delete', index: number, operationId: string): PptxNativeMutationRequestV1 {
  const target = tableTarget(deck, key)?.element;
  if (!target?.source) throw new Error('This table is preserved but cannot be edited.');
  return { expectedSourceRevision: revision(deck), operations: [{ operationId, kind: `table.${axis}.${action}`, elementId: target.id, expectedFingerprintSha256: target.source.fingerprintSha256, index }] };
}

export function slideBackgroundCommand(deck: NativePptxDeck, slideIndex: number, fill: string, operationId: string): PptxNativeMutationRequestV1 {
  const slide = deck.slides[slideIndex]; if (!slide?.source) throw new Error('This slide has no source background anchor.');
  return { expectedSourceRevision: revision(deck), operations: [{ operationId, kind: 'slide.background.set', slideId: slide.id, expectedFingerprintSha256: slide.source.fingerprintSha256, fill }] };
}

export function rotationTarget(deck: NativePptxDeck, key: string) {
  for (const [slideIndex, slide] of deck.slides.entries()) for (const element of slide.elements) if (elementKey(element) === key && (element.kind === 'shape' || element.kind === 'picture' || element.kind === 'text') && element.compatibility.status === 'editable' && element.source) return { element, slideIndex };
}
export function rotationCommand(deck: NativePptxDeck, key: string, rotation60000: number, operationId: string): PptxNativeMutationRequestV1 {
  const target = rotationTarget(deck, key)?.element; if (!target?.source) throw new Error('Rotation is available only for top-level exact shapes, pictures or text boxes.');
  if (!Number.isSafeInteger(rotation60000) || rotation60000 < 0 || rotation60000 >= 21600000) throw new Error('Use an angle from 0 to less than 360 degrees.');
  return { expectedSourceRevision: revision(deck), operations: [{ operationId, kind: target.kind === 'picture' ? 'picture.rotate' : target.kind === 'text' ? 'text.rotate' : 'autoshape.rotate', elementId: target.id, expectedFingerprintSha256: target.source.fingerprintSha256, rotation60000 }] };
}
export function shapeRotationDegrees(element: NativeElement): number { return (element.kind === 'shape' || element.kind === 'picture' || element.kind === 'text') ? (element.rotation60000 ?? 0) / 60000 : 0; }
