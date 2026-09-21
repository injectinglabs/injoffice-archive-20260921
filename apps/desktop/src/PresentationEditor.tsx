import { EditorStatus } from './EditorStatus';
import SlideArrangePanel from './SlideArrangePanel';
import ContextMenu, { presentationContextMenu, selectObjectAt, useContextMenu } from './ContextMenu';
import { arrangeCommand, arrangeTargets, toggleArrangeSelection, type ArrangeAction } from './presentationArrange';
import ShapeArt from './ShapeArt';
import PresentationTextToolbar from './PresentationTextToolbar';
import Ribbon, { RibbonButton, type RibbonTabSpec } from './Ribbon';
import './ribbon.css';
import { useEffect, useRef, useState, type CSSProperties } from 'react';
import { createPptxWasmClient, type PptxNativeExactAutoShapeV1, type PptxNativeExactParagraphV1, type PptxNativeMutationRequestV1 } from '@injoffice/pptx-wasm';
import type { NativeElement, NativePptxDeck, NativeSlide, NativeTransform } from '@injoffice/pptx-native';
import { EMU_PER_INCH, EMU_PER_PIXEL, elementKey, positionElements, shapeCommand, shapeTarget, structureCommand, textCommand, textTarget, recoveryDraft, restorePresentationDraft, insertCommand, deleteElementCommand, transformCommand, pointerTransform, geometryTarget, pictureInsertCommand, tableTarget, rotationCommand, rotationTarget, shapeRotationDegrees, slideBackgroundCommand, tableTopologyCommand, tableInsertCommand, tableCellCommand, tableTextParagraphs, tableSegmentParagraphs } from './presentationCommands';
import './presentation-editor.css';
import PresentationPlayer from './PresentationPlayer';
import { exportNativePptxSlideSvg } from '@injoffice/pptx-render';
import { startPresentationMode, type PresentationModeState } from './presentationMode';
import { engineErrorMessage } from './engine-result';

export interface OfficeEditorProps { registerHistory?: (commands: { undo(): void; redo(): void; canUndo?: boolean; canRedo?: boolean }) => void; registerCommit?: (commit: () => Promise<boolean>) => void; initialRecoveryDraft?: unknown; onRecoveryDraftChange?: (draft: unknown | null) => void; name: string; bytes: Uint8Array; onChange: (bytes: Uint8Array) => void; onBusyChange?: (busy: boolean) => void; onDraftChange?: (dirty: boolean) => void; viewOptions?: { zoom: number; navigation: boolean; focus: boolean } }
type Snapshot = { bytes: Uint8Array; deck: NativePptxDeck };
type Draft = import('./presentationCommands').PresentationDraft;
type PresentationEditorProps = OfficeEditorProps & { initialRecoveryDraft?: unknown; onRecoveryDraftChange?(draft: unknown | null): void; registerCommit?(commit: () => Promise<boolean>): void; registerHistory?(commands: { undo(): void; redo(): void }): void };
const newId = () => `slides-${crypto.randomUUID()}`;
/** PowerPoint's Design gallery, cut down to the backgrounds the native transaction can set. */
const backgroundPresets: [string, string][] = [['FFFFFF', 'White'], ['F5F7FA', 'Light grey'], ['202B3C', 'Dark'], ['2459AD', 'Accent']];
const describeError = (error: unknown) => engineErrorMessage(error);

export default function PresentationEditor({ name, bytes, onChange, onBusyChange, onDraftChange, viewOptions, initialRecoveryDraft, onRecoveryDraftChange, registerCommit, registerHistory }: PresentationEditorProps) {
  const [presentation, setPresentation] = useState<PresentationModeState>();
  const presenting = useRef(false);
  const historyRef = useRef<(direction: 'undo' | 'redo') => void>(() => {});
  const commitRef = useRef<() => Promise<boolean>>(async () => false);
  const composing = useRef(false); const dragging = useRef(false);
  useEffect(() => { registerCommit?.(() => commitRef.current()); }, [registerCommit]);
  const [snapshot, setSnapshot] = useState<Snapshot>();
  const current = useRef<Snapshot | undefined>(undefined);
  const client = useRef<ReturnType<typeof createPptxWasmClient> | undefined>(undefined);
  const callbacks = useRef({ onChange, onBusyChange, onDraftChange, onRecoveryDraftChange }); callbacks.current = { onChange, onBusyChange, onDraftChange, onRecoveryDraftChange };
  const [busy, setBusy] = useState(true); const busyRef = useRef(true);
  const [draft, setDraft] = useState<Draft>(); const draftRef = useRef<Draft | undefined>(undefined);
  const [selected, setSelected] = useState('');
  const [arrangeKeys, setArrangeKeys] = useState<string[]>([]);
  const [index, setIndex] = useState(0);
  const [panel, setPanel] = useState<'shape' | 'position' | 'table' | 'rotation'>('position');
  const [cellSelection, setCellSelection] = useState({ row: 0, column: 0 });
  const [cellSegment, setCellSegment] = useState({ paragraph: 0, run: 0 });
  const [cellTextEdit, setCellTextEdit] = useState<{ text: string; paragraphs: readonly PptxNativeExactParagraphV1[] }>();
  const [tableSize, setTableSize] = useState({ rows: 3, columns: 3 });
  const [tableInsertOpen, setTableInsertOpen] = useState(false);
  const [segment, setSegment] = useState({ paragraph: 0, run: 0 });
  // The run whose text carries the caret on the slide, with the point the caret was placed from.
  const [editing, setEditing] = useState<{ paragraph: number; run: number; offset?: number }>();
  const editingRef = useRef(false); editingRef.current = !!editing;
  const closingEdit = useRef(false);
  const [caretAnchor, setCaretAnchor] = useState<{ left: number; top: number; below: boolean }>();
  const [error, setError] = useState('');
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [ribbonTab, setRibbonTab] = useState('Home');
  // Office's Format pane: closed until an object is selected or Format is pressed.
  const [paneOpen, setPaneOpen] = useState(false);
  const deleteTrigger = useRef<HTMLButtonElement>(null);
  const [undo, setUndo] = useState<Snapshot[]>([]); const [redo, setRedo] = useState<Snapshot[]>([]);
  const undoRef = useRef<Snapshot[]>([]); const redoRef = useRef<Snapshot[]>([]);
  const mounted = useRef(false);
  const workspace = useRef<HTMLDivElement>(null); const [workspaceWidth, setWorkspaceWidth] = useState(740); const [workspaceHeight, setWorkspaceHeight] = useState(600);
  const menu = useContextMenu();
  function markBusy(value: boolean) { busyRef.current = value; setBusy(value); callbacks.current.onBusyChange?.(value); }
  function updateDraft(value: Draft | undefined) { if (!value) setCellTextEdit(undefined); draftRef.current = value; setDraft(value); callbacks.current.onDraftChange?.(!!value); callbacks.current.onBusyChange?.(busyRef.current);
    try { callbacks.current.onRecoveryDraftChange?.(value && current.current ? recoveryDraft(current.current.deck, selected, value) : null); }
    catch (reason) { setError(describeError(reason)); }
  }
  useEffect(() => {
    mounted.current = true; let cancelled = false;
    const worker = createPptxWasmClient({ maxMutationPayloadBytes: 3 * 1024 * 1024 }); client.current = worker; markBusy(true);
    void worker.extract(bytes).then(deck => { if (!cancelled) { const value = { bytes: bytes.slice(), deck }; current.current = value; setSnapshot(value);
      try { const recovered = restorePresentationDraft(deck, initialRecoveryDraft); if (recovered) {
        const target = recovered.draft.kind === 'rotation' ? rotationTarget(deck, recovered.key) : recovered.draft.kind === 'table' ? tableTarget(deck, recovered.key) : recovered.draft.kind === 'geometry' ? geometryTarget(deck, recovered.key) : recovered.draft.kind === 'text' ? textTarget(deck, recovered.key) : shapeTarget(deck, recovered.key);
        if (recovered.draft.kind === 'table') setCellSelection({ row: recovered.draft.row, column: recovered.draft.column });
        setIndex(target!.slideIndex); setSelected(recovered.key); setPanel(recovered.draft.kind === 'geometry' || recovered.draft.kind === 'text' ? (shapeTarget(deck, recovered.key) ? 'shape' : 'position') : recovered.draft.kind);
        draftRef.current = recovered.draft; setDraft(recovered.draft); callbacks.current.onDraftChange?.(true); callbacks.current.onBusyChange?.(busyRef.current);
      } } catch (reason) { setError(describeError(reason)); }
    } })
      .catch(reason => { if (!cancelled) setError(describeError(reason)); })
      .finally(() => { if (!cancelled) markBusy(false); });
    return () => { cancelled = true; mounted.current = false; worker.terminate(); callbacks.current.onBusyChange?.(false); callbacks.current.onDraftChange?.(false); };
  }, []);
  useEffect(() => {
    if (!workspace.current) return;
    const observer = new ResizeObserver(entries => { setWorkspaceWidth(entries[0]?.contentRect.width ?? 740); setWorkspaceHeight(entries[0]?.contentRect.height ?? 600); }); observer.observe(workspace.current);
    return () => observer.disconnect();
  }, [!!snapshot, viewOptions?.focus]);
  // Keep the floating text toolbar over the shape that holds the caret.
  useEffect(() => {
    if (!editing) { setCaretAnchor(undefined); return; }
    const node = workspace.current; if (!node || typeof node.getBoundingClientRect !== 'function') return;
    const place = () => {
      const object = node.querySelector('.presentation-object.is-selected');
      const box = object?.getBoundingClientRect(); const frame = node.getBoundingClientRect();
      if (!box) return;
      const below = box.top - frame.top < 48;
      setCaretAnchor({ left: Math.max(frame.left + 8, Math.min(box.left, frame.right - 320)), top: below ? box.bottom + 8 : box.top - 8, below });
    };
    place();
    node.addEventListener('scroll', place); window.addEventListener('resize', place);
    return () => { node.removeEventListener('scroll', place); window.removeEventListener('resize', place); };
  }, [editing, index, selected, workspaceWidth, viewOptions?.zoom]);
  const blocked = busy || !!draft || confirmDelete;
  const canUndo = !blocked && undo.length > 0, canRedo = !blocked && redo.length > 0;
  useEffect(() => { registerHistory?.({ undo: () => historyRef.current('undo'), redo: () => historyRef.current('redo'), canUndo, canRedo }); }, [registerHistory, canUndo, canRedo]);
  const slide = snapshot?.deck.slides[index];
  // What the canvas used to caption and hide behind a disclosure, in one tooltip.
  const previewNote = ['Positioned preview: text wrapping may differ in PowerPoint.',
    'Exact text and supported shapes can be edited. Images use embedded previews when available. Unsupported content stays in the file and appears as a placeholder. Slide commands can be refused for notes, comments, links, sections or other relationships that cannot be changed safely.',
    ...(slide?.compatibility.diagnostics ?? []).slice(0, 5).map(diagnostic => diagnostic.message)].join('\n');
  const selectedItem = slide && positionElements(slide.elements).find(item => elementKey(item.element) === selected);
  const text = snapshot && textTarget(snapshot.deck, selected);
  const shape = snapshot && shapeTarget(snapshot.deck, selected);
  const table = snapshot && tableTarget(snapshot.deck, selected);
  const tableCell = table?.element.table.rows[cellSelection.row]?.[cellSelection.column];
  const cellDraft = draft?.kind === 'table' ? draft : undefined;
  const paragraphs = draft?.kind === 'text' ? draft.paragraphs : text?.paragraphs;
  const selectedParagraph = paragraphs?.[segment.paragraph]; const run = selectedParagraph?.runs[segment.run];
  const shapeValue = draft?.kind === 'shape' ? draft.shape : shape?.autoShape;
  const hasSelection = !!selectedItem || arrangeKeys.length > 1;
  // Selecting an object opens the pane; clearing the selection closes it again.
  // A manual toggle in between is kept until the selection changes.
  useEffect(() => { setPaneOpen(hasSelection); }, [hasSelection]);
  const arrangeable = snapshot ? arrangeTargets(snapshot.deck, index) : [];
  function choose(key: string, additive = false) {
    if (busyRef.current || draftRef.current || confirmDelete) return;
    if (additive && current.current) {
      const allowed = arrangeTargets(current.current.deck,index).map(elementKey);
      if (!allowed.includes(key)) return;
      const next = toggleArrangeSelection(arrangeKeys,key);
      if (next.length === 1) choose(next[0]!);
      else { setArrangeKeys(next); setSelected(''); }
      return;
    }
    setArrangeKeys(current.current && arrangeTargets(current.current.deck,index).some(item => elementKey(item) === key) ? [key] : []);
    setSelected(key); setCellSelection({ row: 0, column: 0 }); setSegment({ paragraph: 0, run: 0 }); setEditing(undefined); setPanel(current.current && tableTarget(current.current.deck, key) ? 'table' : current.current && shapeTarget(current.current.deck, key) ? 'shape' : current.current && geometryTarget(current.current.deck, key) ? 'position' : 'shape'); setError('');
  }
  /** Put the caret in a run painted on the slide. The ribbon and the floating toolbar keep working on the same segment. */
  function beginEdit(key: string, at: { paragraph: number; run: number; offset?: number }) {
    if (busyRef.current || confirmDelete || !current.current) return;
    const target = textTarget(current.current.deck, key); if (!target) return;
    const paragraph = Math.min(Math.max(at.paragraph, 0), target.paragraphs.length - 1);
    const run = Math.min(Math.max(at.run, 0), (target.paragraphs[paragraph]?.runs.length ?? 1) - 1);
    if (run < 0) return;
    closingEdit.current = false; setSegment({ paragraph, run }); setEditing({ paragraph, run, offset: at.offset });
  }
  /** Canvas selection. A pending caret edit is committed first so one click moves on. */
  async function selectFromCanvas(key: string, additive: boolean, at?: { paragraph: number; run: number; offset?: number }) {
    if (editingRef.current && draftRef.current?.kind === 'text') {
      closingEdit.current = true; setEditing(undefined);
      if (!await applyDraft()) return;
    }
    choose(key, additive);
    if (!additive && at) beginEdit(key, at);
  }
  function editText(value: string) {
    if (!editingRef.current) return;
    // The native text transaction has no line breaks inside a run; pasted ones become spaces.
    textPatch({ text: value.replace(/[\r\n\t]+/g, ' ') });
  }
  async function commitEdit() {
    closingEdit.current = true; setEditing(undefined);
    if (draftRef.current) await applyDraft();
  }
  function cancelEdit() {
    closingEdit.current = true; setEditing(undefined);
    if (draftRef.current?.kind === 'text') { updateDraft(undefined); setError(''); }
  }
  function arrange(action: ArrangeAction) {
    if (blocked || !current.current || dragging.current || composing.current) return;
    try { void applyCommand(arrangeCommand(current.current.deck,index,arrangeKeys,action,newId())); }
    catch (reason) { setError(describeError(reason)); }
  }
  function selectSlide(next: number) { if (blocked) return; setIndex(next); setSelected(''); setArrangeKeys([]); setSegment({ paragraph: 0, run: 0 }); setEditing(undefined); setError(''); }
  function textPatch(patch: Partial<NonNullable<typeof run>>, alignment?: 'left' | 'center' | 'right') {
    if (!paragraphs || busyRef.current) return;
    const next = structuredClone(paragraphs) as PptxNativeExactParagraphV1[];
    const p = next[segment.paragraph]; if (!p?.runs[segment.run]) return;
    next[segment.paragraph] = alignment ? { ...p, align: alignment } : { ...p, runs: p.runs.map((value, i) => i === segment.run ? { ...value, ...patch } : value) };
    updateDraft(JSON.stringify(next) === JSON.stringify(text?.paragraphs) ? undefined : { kind: 'text', paragraphs: next });
  }
  function shapePatch(patch: Partial<PptxNativeExactAutoShapeV1>) {
    if (!shapeValue || busyRef.current) return;
    const next = { ...structuredClone(shapeValue), ...patch };
    updateDraft(JSON.stringify(next) === JSON.stringify(shape?.autoShape) ? undefined : { kind: 'shape', shape: next });
  }
  function geometryPatch(key: string, transform: NativeTransform) {
    if (busyRef.current || !current.current) return;
    const shape = shapeTarget(current.current.deck, key);
    if (shape) { const next = { ...shape.autoShape, transform }; updateDraft(JSON.stringify(next) === JSON.stringify(shape.autoShape) ? undefined : { kind: 'shape', shape: next }); setPanel('shape'); }
    else {
      const element = current.current.deck.slides.flatMap(s => positionElements(s.elements)).find(item => elementKey(item.element) === key)?.element;
      if (!element || !geometryTarget(current.current.deck, key)) return;
      updateDraft(JSON.stringify(transform) === JSON.stringify(element.transform) ? undefined : { kind: 'geometry', transform }); setPanel('position');
    }
  }
  function install(value: Snapshot) { current.current = value; setSnapshot(value); callbacks.current.onChange(value.bytes.slice()); }
  async function applyCommand(request: PptxNativeMutationRequestV1, nextIndex = index) {
    const previous = current.current; const worker = client.current;
    if (!previous || !worker || busyRef.current) return false;
    markBusy(true); setError('');
    try {
      const output = await worker.apply(previous.bytes, previous.deck, request);
      const deck = await worker.extract(output);
      if (!mounted.current) return false;
      if (deck.sourceRevision === previous.deck.sourceRevision) throw new Error('The change did not produce a new presentation.');
      let history = [...undoRef.current, previous].slice(-20);
      while (history.length > 1 && history.reduce((sum, value) => sum + value.bytes.byteLength, 0) > 128 * 1024 * 1024) history = history.slice(1);
      undoRef.current = history; redoRef.current = []; setUndo(history); setRedo([]);
      install({ bytes: output, deck }); updateDraft(undefined);
      setIndex(Math.max(0, Math.min(nextIndex, deck.slides.length - 1)));
      if ((request.operations[0]?.kind.startsWith('slide.') && request.operations[0]?.kind !== 'slide.background.set') || request.operations[0]?.kind === 'element.delete') { setSelected(''); setArrangeKeys([]); }
      if (request.operations[0]?.kind === 'table.insert' || request.operations[0]?.kind === 'picture.insert' || request.operations[0]?.kind === 'text.insert' || request.operations[0]?.kind === 'autoshape.insert') {
        const item = deck.slides[nextIndex]?.elements.at(-1); if (item) { setArrangeKeys([]); setSelected(elementKey(item)); setPanel(request.operations[0].kind === 'table.insert' ? 'table' : request.operations[0].kind === 'autoshape.insert' ? 'shape' : 'position'); setSegment({ paragraph: 0, run: 0 }); setCellSelection({ row: 0, column: 0 }); }
      }
      return true;
    } catch (reason) { if (mounted.current) setError(describeError(reason)); return false; }
    finally { if (mounted.current) markBusy(false); }
  }
  async function applyDraft(): Promise<boolean> {
    const value = current.current; const pending = draftRef.current;
    if (!value || presenting.current || busyRef.current || composing.current || dragging.current) return false;
    if (!pending) return true;
    try { return await applyCommand(pending.kind === 'rotation' ? rotationCommand(value.deck, selected, pending.rotation60000, newId()) : pending.kind === 'table' ? tableCellCommand(value.deck, selected, pending, newId()) : pending.kind === 'geometry' ? transformCommand(value.deck, selected, pending.transform, newId()) : pending.kind === 'text' ? textCommand(value.deck, selected, pending.paragraphs, newId()) : shapeCommand(value.deck, selected, pending.shape, newId())); }
    catch (reason) { setError(describeError(reason)); return false; }
  }
  commitRef.current = applyDraft;
  function insert(kind: 'slide' | 'text' | 'shape') {
    const value = current.current; if (!value || blocked) return;
    try { void applyCommand(insertCommand(value.deck, index, kind, newId()), kind === 'slide' ? index + 1 : index); } catch (reason) { setError(describeError(reason)); }
  }
  function insertTable() {
    if (!current.current || blocked) return;
    try { const request = tableInsertCommand(current.current.deck, index, tableSize.rows, tableSize.columns, newId()); setTableInsertOpen(false); void applyCommand(request); }
    catch (reason) { setError(describeError(reason)); }
  }
  async function changeTableTrack(axis: 'row' | 'column', action: 'insert' | 'delete') {
    const value = current.current; if (!value || blocked || busyRef.current || draftRef.current) return;
    try { const at = cellSelection[axis] + (action === 'insert' ? 1 : 0);
      if (await applyCommand(tableTopologyCommand(value.deck, selected, axis, action, at, newId()))) {
        const next = tableTarget(current.current!.deck, selected)?.element.table;
        if (next) setCellSelection(previous => ({ row: Math.min(previous.row, next.rows.length - 1), column: Math.min(previous.column, next.columnWidths.length - 1), [axis]: action === 'insert' ? at : Math.min(at, (axis === 'row' ? next.rows.length : next.columnWidths.length) - 1) }));
      }
    } catch (reason) { setError(describeError(reason)); }
  }
  async function exportSvg() {
    const value = current.current; if (!value || blocked || busyRef.current || draftRef.current || presenting.current) return;
    const write = (window.injDesktop as typeof window.injDesktop & { exportBytes?(input: { name: string; bytes: Uint8Array }): Promise<{ name: string } | null> })?.exportBytes;
    if (!write) { setError('SVG export is unavailable in this app build.'); return; }
    markBusy(true); setError('');
    try {
      const svg = await exportNativePptxSlideSvg(value.deck, index, { validateImage: async (bytes, contentType) => {
        if (typeof createImageBitmap !== 'function') throw new Error('Local image decoding is unavailable.');
        const bitmap = await createImageBitmap(new Blob([bytes.slice().buffer as ArrayBuffer], { type: contentType }), { imageOrientation: 'none' });
        try { return { width: bitmap.width, height: bitmap.height }; } finally { bitmap.close(); }
      } });
      if (!mounted.current) return;
      await write({ name: `${name.replace(/\.pptx$/i, '')}-slide-${index + 1}.svg`, bytes: new TextEncoder().encode(svg) });
    } catch (reason) { if (mounted.current) setError(describeError(reason)); }
    finally { if (mounted.current) markBusy(false); }
  }
  function present(from: number) {
    if (!current.current || busyRef.current || confirmDelete || dragging.current || composing.current) return;
    presenting.current = true; setPresentation(startPresentationMode(current.current.deck, from, !!draftRef.current));
  }
  function changeBackground(fill: string) {
    const value = current.current; if (!value || blocked || busyRef.current || draftRef.current) return;
    try { void applyCommand(slideBackgroundCommand(value.deck, index, fill, newId())); } catch (reason) { setError(describeError(reason)); }
  }
  function patchCell(textValue?: string, fillValue?: { fill: string | undefined }) {
    if (!tableCell || busyRef.current) return;
    const original = tableCell.paragraphs as PptxNativeExactParagraphV1[];
    const paragraphs = textValue === undefined ? (cellDraft?.paragraphs ?? structuredClone(original)) : original.length === 1 && original[0]?.runs.length === 1 && textValue === tableCell.text ? structuredClone(original) : original.length === 1 && original[0]?.runs.length === 1 ? tableTextParagraphs(original, textValue) : tableSegmentParagraphs(cellTextEdit?.paragraphs ?? cellDraft?.paragraphs ?? original, cellSegment.paragraph, cellSegment.run, textValue);
    if (textValue !== undefined && (original.length > 1 || original[0]?.runs.length !== 1)) setCellTextEdit({ text: textValue, paragraphs: cellTextEdit?.paragraphs ?? cellDraft?.paragraphs ?? original });
    const fill = fillValue ? fillValue.fill : cellDraft ? cellDraft.fill : tableCell.fill;
    updateDraft(JSON.stringify(paragraphs) === JSON.stringify(original) && fill === tableCell.fill ? undefined : { kind: 'table', row: cellSelection.row, column: cellSelection.column, paragraphs, ...(fill === undefined ? {} : { fill }) });
  }
  function chooseCell(key: string, row: number, column: number) { if (blocked) return; choose(key); setCellSelection({ row, column }); setCellSegment({ paragraph: 0, run: 0 }); setCellTextEdit(undefined); setPanel('table'); }
  async function insertPicture() {
    const value = current.current; if (!value || blocked) return;
    const pick = (window.injDesktop as typeof window.injDesktop & { pickAsset?(kind: 'image'): Promise<{ name: string; bytes: Uint8Array } | null> })?.pickAsset;
    if (!pick) { setError('Image selection is unavailable in this app build.'); return; }
    markBusy(true); setError('');
    try { const asset = await pick('image'); if (!mounted.current || !asset) return;
      const request = pictureInsertCommand(value.deck, index, asset.bytes, newId());
      markBusy(false); await applyCommand(request);
    } catch (reason) { if (mounted.current) setError(describeError(reason)); }
    finally { if (mounted.current) markBusy(false); }
  }
  function deleteObject() {
    const value = current.current; if (!value || blocked) return;
    try { void applyCommand(deleteElementCommand(value.deck, selected, newId())); } catch (reason) { setError(describeError(reason)); }
  }
  function structure(action: 'duplicate' | 'delete' | 'previous' | 'next') {
    const value = current.current; if (!value || busyRef.current || draftRef.current) return;
    try { const next = action === 'duplicate' || action === 'next' ? index + 1 : action === 'previous' ? index - 1 : Math.min(index, value.deck.slides.length - 2); void applyCommand(structureCommand(value.deck, index, action, newId()), next); }
    catch (reason) { setError(describeError(reason)); }
  }
  function travel(direction: 'undo' | 'redo') {
    if (presenting.current || blocked || busyRef.current || draftRef.current || composing.current || dragging.current || !current.current) return;
    const source = direction === 'undo' ? undoRef : redoRef; const destination = direction === 'undo' ? redoRef : undoRef;
    const next = source.current.at(-1); if (!next) return;
    source.current = source.current.slice(0, -1); destination.current = [...destination.current, current.current];
    setUndo(undoRef.current); setRedo(redoRef.current); install(next); setIndex(Math.min(index, next.deck.slides.length - 1)); setSelected(''); setArrangeKeys([]); setError('');
  }
  historyRef.current = travel;
  const scale = snapshot ? Math.max(.1, Math.min(1, (workspaceWidth - 64) / (snapshot.deck.size.cx / EMU_PER_PIXEL), (workspaceHeight - 34) / (snapshot.deck.size.cy / EMU_PER_PIXEL))) * Math.max(.5, Math.min(2, (viewOptions?.zoom ?? 100) / 100)) : 1;
  // PowerPoint's ribbon, applied to the controls this editor already has.
  const ribbonTabs: RibbonTabSpec[] = [
    { id: 'File', label: 'File', groups: [{ id: 'export', label: 'Export', children: <RibbonButton icon="svg" label="Export slide SVG" title="Export the current supported slide. Text and unsupported objects require the original PPTX." disabled={blocked || !snapshot} onClick={() => void exportSvg()} /> }] },
    { id: 'Home', label: 'Home', groups: [
      { id: 'slides', label: 'Slides', children: <>
        <RibbonButton icon="newSlide" label="New slide" disabled={blocked || !slide} onClick={() => insert('slide')} />
        <RibbonButton icon="duplicate" label="Duplicate slide" disabled={blocked || !slide} onClick={() => structure('duplicate')} />
        <RibbonButton icon="moveEarlier" label="Move earlier" disabled={blocked || !slide || index === 0} onClick={() => structure('previous')} />
        <RibbonButton icon="moveLater" label="Move later" disabled={blocked || !snapshot || index >= snapshot.deck.slides.length - 1} onClick={() => structure('next')} />
        <RibbonButton icon="deleteSlide" label="Delete slide" disabled={blocked || !snapshot || snapshot.deck.slides.length <= 1} ref={deleteTrigger} onClick={() => setConfirmDelete(true)} />
      </> },
      { id: 'font', label: 'Font', children: <PresentationTextToolbar section="font" run={run} align={selectedParagraph?.align} disabled={busy || confirmDelete || !text || (!!draft && draft.kind !== 'text')} onRunChange={patch => textPatch(patch)} onAlignChange={align => textPatch({}, align)} /> },
      { id: 'paragraph', label: 'Paragraph', children: <PresentationTextToolbar section="paragraph" run={run} align={selectedParagraph?.align} disabled={busy || confirmDelete || !text || (!!draft && draft.kind !== 'text')} onRunChange={patch => textPatch(patch)} onAlignChange={align => textPatch({}, align)} /> },
      { id: 'drawing', label: 'Drawing', children: <>
        <RibbonButton icon="align" label="Arrange" title="Align or space the objects selected on this slide" disabled={blocked || arrangeable.length < 2} aria-expanded={paneOpen} onClick={() => setPaneOpen(true)} />
        <RibbonButton icon="sidebar" label="Format" title="Show or hide the Format pane for the selected object" disabled={!snapshot} aria-pressed={paneOpen} onClick={() => setPaneOpen(!paneOpen)} />
        <RibbonButton icon="deleteObject" label="Delete object" disabled={blocked || !selectedItem || selectedItem.grouped || (!text && !shape && !geometryTarget(snapshot!.deck, selected))} onClick={deleteObject} />
      </> },
    ] },
    { id: 'Insert', label: 'Insert', groups: [
      { id: 'tables', label: 'Tables', children: <RibbonButton icon="table" label="Table" disabled={blocked || !slide} aria-expanded={tableInsertOpen} onClick={() => setTableInsertOpen(!tableInsertOpen)} /> },
      { id: 'images', label: 'Images', children: <RibbonButton icon="image" label="Picture" disabled={blocked || !slide} onClick={() => void insertPicture()} /> },
      { id: 'illustrations', label: 'Illustrations', children: <RibbonButton icon="shape" label="Shape" disabled={blocked || !slide} onClick={() => insert('shape')} /> },
      { id: 'text', label: 'Text', children: <RibbonButton icon="textbox" label="Text box" disabled={blocked || !slide} onClick={() => insert('text')} /> },
    ] },
    { id: 'Design', label: 'Design', groups: [{ id: 'customize', label: 'Customize', children: <div className="presentation-background-gallery" role="toolbar" aria-label="Slide background">
      {backgroundPresets.map(([color, label]) => <button key={color} type="button" className="ribbon-button presentation-background-swatch" title={`${label} background`} aria-label={`${label} background`} aria-pressed={(slide?.background ?? '') === color}
        disabled={blocked || !slide || slide.compatibility.diagnostics.some(d => d.code === 'pptx.unsupported-background')} onClick={() => changeBackground(color)}>
        <svg viewBox="0 0 32 20" aria-hidden="true"><rect x=".5" y=".5" width="31" height="19" rx="1.5" fill={`#${color}`} stroke="currentColor" strokeOpacity=".35" /></svg>
        <span className="ribbon-button-label">{label}</span>
      </button>)}
    </div> }] },
    { id: 'SlideShow', label: 'Slide Show', groups: [{ id: 'start', label: 'Start Slide Show', children: <>
      <RibbonButton icon="present" label="From Beginning" title="Start the slide show from the first slide" disabled={busy || !snapshot || confirmDelete} onClick={() => present(0)} />
      <RibbonButton icon="present" label="From Current Slide" title="Start the slide show from the slide on screen" disabled={busy || !snapshot || confirmDelete} onClick={() => present(index)} />
    </> }] },
  ];
  return <div className="presentation-editor" aria-label="Presentation editor" aria-busy={busy} onKeyDown={event => {
      if (!(event.ctrlKey || event.metaKey) || event.altKey || blocked || composing.current || dragging.current) return;
      const key = event.key.toLowerCase(); if (key !== 'z' && key !== 'y') return;
      event.preventDefault(); event.stopPropagation(); travel(key === 'y' || event.shiftKey ? 'redo' : 'undo');
    }} onCompositionStart={() => { composing.current = true; }} onCompositionEnd={() => { composing.current = false; }}>
    {presentation && <PresentationPlayer initial={presentation} onExit={() => { presenting.current = false; setPresentation(undefined); }} renderSlide={(slide, scale) => <SlideCanvas deck={presentation.deck} slide={slide} scale={scale} thumbnail />} />}
    <Ribbon label="Presentation tools" tabs={ribbonTabs} active={ribbonTab} onChange={setRibbonTab} />
    {(tableInsertOpen || (draft && !editing)) && <div className="presentation-toolbar">
      {tableInsertOpen && <div className="presentation-table-options" aria-label="Insert table"><label>Rows <input type="number" min="1" max="100" aria-label="New table rows" value={Number.isFinite(tableSize.rows) ? tableSize.rows : ''} onChange={event => setTableSize({ ...tableSize, rows: event.target.valueAsNumber })} /></label><label>Columns <input type="number" min="1" max="100" aria-label="New table columns" value={Number.isFinite(tableSize.columns) ? tableSize.columns : ''} onChange={event => setTableSize({ ...tableSize, columns: event.target.valueAsNumber })} /></label><RibbonButton icon="check" label="Insert table" disabled={blocked} onClick={insertTable} /><RibbonButton icon="close" label="Cancel" onClick={() => setTableInsertOpen(false)} /></div>}
      {draft && <div className="presentation-pending"><span>Pending changes</span><RibbonButton className="presentation-primary ribbon-primary" icon="check" label="Apply changes" disabled={busy} onClick={applyDraft} /><RibbonButton icon="close" label="Cancel" disabled={busy} onClick={() => { updateDraft(undefined); setError(''); }} /></div>}
    </div>}
    {error && <div className="presentation-error" role="alert"><span>{error}</span><button aria-label="Dismiss presentation error" onClick={() => setError('')}>×</button></div>}
    {!snapshot ? <div className="presentation-loading" role="status">{busy ? 'Opening presentation…' : 'This presentation could not be opened.'}</div> : <div className="presentation-layout">
      {!viewOptions?.focus && <nav className="presentation-thumbnails" aria-label="Slides"><div className="presentation-rail-title">Slides <span>{snapshot.deck.slides.length}</span></div>{snapshot.deck.slides.map((item, i) => <SlideThumbnail key={item.id} deck={snapshot.deck} slide={item} number={i + 1} current={index === i} disabled={blocked} onSelect={() => selectSlide(i)} />)}</nav>}
      <div className="presentation-workspace" ref={workspace} onContextMenu={event => { selectObjectAt(event); menu.open(event); }}>
        <div className="presentation-canvas-scroll">{slide && <SlideCanvas deck={snapshot.deck} slide={slide} scale={scale} selected={selected} selectedKeys={arrangeKeys} onSelect={(key, additive, at) => void selectFromCanvas(key, !!additive, at)} disabled={busy || confirmDelete} editing={editing} onTextInput={editText} onTextCommit={() => void commitEdit()} onTextCancel={cancelEdit} onTextBlur={keepEditing => { if (closingEdit.current) { closingEdit.current = false; return; } if (!keepEditing) void commitEdit(); }} draft={draft} onGeometry={geometryPatch} selectedCell={cellSelection} onCellSelect={chooseCell} onGestureChange={value => { dragging.current = value; }} />}</div>
        {editing && text && caretAnchor && <div className="presentation-floating-toolbar" style={{ left: caretAnchor.left, top: caretAnchor.top, transform: caretAnchor.below ? undefined : 'translateY(-100%)' }}>
          <PresentationTextToolbar run={run} align={selectedParagraph?.align} disabled={busy || confirmDelete} onRunChange={patch => textPatch(patch)} onAlignChange={align => textPatch({}, align)} />
        </div>}
      </div>
      {paneOpen && <aside className="presentation-inspector" aria-label="Format">
        <div className="presentation-inspector-header">
          <h2>{selectedItem?.element.name || (selectedItem ? selectedItem.element.kind : arrangeKeys.length > 1 ? `${arrangeKeys.length} objects selected` : 'Format')}</h2>
          <button className="presentation-pane-close" aria-label="Close the Format pane" title="Close the Format pane" onClick={() => setPaneOpen(false)}>×</button>
        </div>
        {arrangeKeys.length > 1 ? <p className="presentation-help">Choose a single object to edit its content or appearance.</p> : <>
        {!selectedItem ? <p className="presentation-help">Choose text or a shape on the slide to edit it.</p> : <>
          <div className="presentation-inspector-tabs">{rotationTarget(snapshot.deck, selected) && <button aria-pressed={panel === 'rotation'} disabled={!!draft || busy} onClick={() => setPanel('rotation')}>Rotation</button>}{table && <button aria-pressed={panel === 'table'} disabled={!!draft || busy} onClick={() => setPanel('table')}>Cell</button>}{snapshot && geometryTarget(snapshot.deck, selected) && <button aria-pressed={panel === 'position'} disabled={!!draft || busy} onClick={() => setPanel('position')}>Position</button>}{shape && <button aria-pressed={panel === 'shape'} disabled={!!draft || busy} onClick={() => setPanel('shape')}>Shape</button>}</div>
          {!text && !shape && !geometryTarget(snapshot.deck, selected) && <p className="presentation-help">This object is preserved in the original file. Editing is not available for its current format.</p>}
          {text && <p className="presentation-help">Click the text on the slide to put the caret in it. The ribbon Font and Paragraph groups, and the toolbar above the slide, format the segment holding the caret.</p>}
          {panel === 'table' && table && tableCell && <fieldset disabled={busy}>
            <div className="presentation-table-tracks" aria-label="Table rows and columns">
              <button disabled={blocked || table.element.table.rows.length >= 100 || (table.element.table.rows.length + 1) * table.element.table.columnWidths.length > 1000} onClick={() => void changeTableTrack('row', 'insert')}>Insert row below</button>
              <button disabled={blocked || table.element.table.rows.length <= 1} onClick={() => void changeTableTrack('row', 'delete')}>Delete row</button>
              <button disabled={blocked || table.element.table.columnWidths.length >= 100 || table.element.table.rows.length * (table.element.table.columnWidths.length + 1) > 1000} onClick={() => void changeTableTrack('column', 'insert')}>Insert column after</button>
              <button disabled={blocked || table.element.table.columnWidths.length <= 1} onClick={() => void changeTableTrack('column', 'delete')}>Delete column</button>
            </div><p className="presentation-inspector-help">New tracks inherit the selected cells’ appearance. The table grows or shrinks with its tracks.</p>
            <div className="presentation-inspector-row"><label>Row<select aria-label="Table cell row" disabled={!!draft} value={cellSelection.row} onChange={event => { setCellSelection({ ...cellSelection, row: Number(event.target.value) }); setCellSegment({ paragraph: 0, run: 0 }); }}>{table.element.table.rows.map((_, i) => <option key={i} value={i}>{i + 1}</option>)}</select></label><label>Column<select aria-label="Table cell column" disabled={!!draft} value={cellSelection.column} onChange={event => { setCellSelection({ ...cellSelection, column: Number(event.target.value) }); setCellSegment({ paragraph: 0, run: 0 }); }}>{table.element.table.columnWidths.map((_, i) => <option key={i} value={i}>{i + 1}</option>)}</select></label></div>
            {(tableCell.paragraphs!.length > 1 || tableCell.paragraphs![0]!.runs.length > 1) && <label>Text segment<select aria-label="Table text segment" value={`${cellSegment.paragraph}:${cellSegment.run}`} onChange={event => { const [paragraph, run] = event.target.value.split(':').map(Number); setCellSegment({ paragraph: paragraph!, run: run! }); setCellTextEdit(undefined); }}>{(cellDraft?.paragraphs ?? tableCell.paragraphs!).flatMap((paragraph, pi) => paragraph.runs.map((run, ri) => <option key={`${pi}:${ri}`} value={`${pi}:${ri}`}>Paragraph {pi + 1}, segment {ri + 1}: {run.text.slice(0, 35) || 'Empty text'}</option>))}</select></label>}
            <label>Cell text<textarea rows={6} aria-label="Table cell text" value={tableCell.paragraphs!.length === 1 && tableCell.paragraphs![0]!.runs.length === 1 ? (cellDraft ? cellDraft.paragraphs.map(p => p.runs.map(r => r.text).join('')).join('\n') : tableCell.text) : cellTextEdit?.text ?? (cellDraft?.paragraphs ?? tableCell.paragraphs!)[cellSegment.paragraph]?.runs[cellSegment.run]?.text ?? ''} onChange={event => patchCell(event.target.value)} /></label>
            <div className="presentation-inspector-row"><label className="presentation-checkbox"><input type="checkbox" aria-label="Table cell has fill" checked={(cellDraft ? cellDraft.fill : tableCell.fill) !== undefined} onChange={event => patchCell(undefined, { fill: event.target.checked ? 'F5F7FA' : undefined })} />Fill</label><input type="color" aria-label="Table cell fill" disabled={!(cellDraft ? cellDraft.fill : tableCell.fill)} value={`#${(cellDraft ? cellDraft.fill : tableCell.fill) ?? 'FFFFFF'}`} onChange={event => patchCell(undefined, { fill: event.target.value.slice(1).toUpperCase() })} /></div>
            <p className="presentation-help">Click a cell to select it. Edits preserve other segments and paragraph formatting. New lines inherit the edited paragraph’s formatting. Dotted borders are editing guides; the saved table has no visible borders.</p>
          </fieldset>}
          {panel === 'position' && geometryTarget(snapshot.deck, selected) && <fieldset disabled={busy}>
            <h3>{selectedItem.grouped ? 'Position in group (inches)' : 'Position on slide (inches)'}</h3>
            <div className="presentation-geometry">{(['x', 'y', 'cx', 'cy'] as const).map((key, i) => { const transform = draft?.kind === 'geometry' ? draft.transform : selectedItem.element.transform; return <label key={key}>{['Left', 'Top', 'Width', 'Height'][i]}<input type="number" step=".05" min={i > 1 ? .01 : undefined} aria-label={`${selectedItem.element.kind === 'picture' ? 'Picture' : selectedItem.element.kind === 'table' ? 'Table' : 'Text box'} ${['left', 'top', 'width', 'height'][i]}`} value={Number.isFinite(transform[key]) ? Math.round(transform[key] / EMU_PER_INCH * 10000) / 10000 : ''} onChange={event => geometryPatch(selected, { ...transform, [key]: Math.round(event.target.valueAsNumber * EMU_PER_INCH) })} /></label>; })}</div>
            <p className="presentation-help">Drag the selected object to move it. Drag its lower-right handle to resize. Apply to keep your changes.</p>
          </fieldset>}
          {panel === 'rotation' && rotationTarget(snapshot.deck, selected) && <fieldset disabled={busy}>
            <label>Rotation (degrees)<input type="number" min="0" max="359.999" step="1" aria-label={`${selectedItem.element.kind === 'picture' ? 'Picture' : selectedItem.element.kind === 'text' ? 'Text box' : 'Shape'} rotation degrees`} value={draft?.kind === 'rotation' ? (Number.isFinite(draft.rotation60000) ? draft.rotation60000 / 60000 : '') : shapeRotationDegrees(selectedItem.element)} onChange={event => { const rotation60000 = Math.round(event.target.valueAsNumber * 60000); updateDraft(rotation60000 === shapeRotationDegrees(selectedItem.element) * 60000 ? undefined : { kind: 'rotation', rotation60000 }); }} /></label>
            <p className="presentation-help">Clockwise around the object’s center. Use Width and Height in the Shape or Position panel to resize; dragging preserves the angle.</p>
          </fieldset>}
          {panel === 'shape' && shapeValue && <fieldset disabled={busy}>
            <label>Shape<select aria-label="Shape preset" value={shapeValue.preset} onChange={event => shapePatch({ preset: event.target.value as PptxNativeExactAutoShapeV1['preset'] })}><option value="rect">Rectangle</option><option value="ellipse">Ellipse</option><option value="triangle">Triangle</option><option value="diamond">Diamond</option></select></label>
            <h3>{selectedItem.grouped ? 'Position in group (inches)' : 'Position on slide (inches)'}</h3><div className="presentation-geometry">{(['x', 'y', 'cx', 'cy'] as const).map((key, i) => <label key={key}>{['Left', 'Top', 'Width', 'Height'][i]}<input type="number" step=".05" min={i > 1 ? .01 : undefined} aria-label={`Shape ${['left', 'top', 'width', 'height'][i]}`} value={Number.isFinite(shapeValue.transform[key]) ? Math.round(shapeValue.transform[key] / EMU_PER_INCH * 10000) / 10000 : ''} onChange={event => shapePatch({ transform: { ...shapeValue.transform, [key]: Math.round(event.target.valueAsNumber * EMU_PER_INCH) } })} /></label>)}</div>
            <div className="presentation-inspector-row"><label className="presentation-checkbox"><input type="checkbox" checked={shapeValue.fill !== undefined} onChange={event => shapePatch({ fill: event.target.checked ? 'DCE8F7' : undefined })} />Fill</label><input type="color" aria-label="Shape fill color" disabled={!shapeValue.fill} value={`#${shapeValue.fill || 'FFFFFF'}`} onChange={event => shapePatch({ fill: event.target.value.slice(1).toUpperCase() })} /></div>
            <label className="presentation-checkbox"><input type="checkbox" checked={!!shapeValue.stroke} onChange={event => shapePatch({ stroke: event.target.checked ? { color: '2459AD', widthEmu: 12700, cap: 'flat', join: 'round', dash: 'solid' } : undefined })} />Outline</label>
            {shapeValue.stroke && <div className="presentation-inspector-row"><label>Width (pt)<input type="number" min=".1" max="100" step=".5" aria-label="Shape outline width" value={Number.isFinite(shapeValue.stroke.widthEmu) ? shapeValue.stroke.widthEmu / 12700 : ''} onChange={event => shapePatch({ stroke: { ...shapeValue.stroke!, widthEmu: Math.round(event.target.valueAsNumber * 12700) } })} /></label><label>Color<input type="color" aria-label="Shape outline color" value={`#${shapeValue.stroke.color}`} onChange={event => shapePatch({ stroke: { ...shapeValue.stroke!, color: event.target.value.slice(1).toUpperCase() } })} /></label></div>}
          </fieldset>}
        </>}
        </>}
        {arrangeable.length > 1 && <SlideArrangePanel elements={arrangeable} keys={arrangeKeys} disabled={blocked} onToggle={key => choose(key,true)} onArrange={arrange} />}
      </aside>}
    </div>}
    {snapshot && <EditorStatus label="Presentation status">
      <span className="presentation-status-slide">Slide {index + 1} of {snapshot.deck.slides.length}</span>
      <span>{editing ? 'Editing text · Enter applies, Esc cancels' : selectedItem ? `Selected: ${selectedItem.element.name || selectedItem.element.kind}` : ''}</span>
      <span className="presentation-status-note" tabIndex={0} role="note" aria-label="Preview note" title={previewNote}>&#9432;</span>
    </EditorStatus>}
    {menu.anchor && snapshot && <ContextMenu anchor={menu.anchor} label="Slide" onClose={menu.close} items={presentationContextMenu({ object: !!selectedItem && !selectedItem.grouped && (!!text || !!shape || !!geometryTarget(snapshot.deck, selected)), slide: !!slide, disabled: blocked, canDeleteSlide: snapshot.deck.slides.length > 1, onDeleteObject: deleteObject, onNewSlide: () => insert('slide'), onDuplicateSlide: () => structure('duplicate'), onDeleteSlide: () => setConfirmDelete(true) })} />}
    {confirmDelete &&<div className="presentation-modal"><section role="alertdialog" aria-modal="true" aria-labelledby="presentation-delete-title" onKeyDown={event => {
      if (event.key === 'Escape') { event.preventDefault(); setConfirmDelete(false); deleteTrigger.current?.focus(); }
      if (event.key === 'Tab') { const buttons = Array.from(event.currentTarget.querySelectorAll('button')); const next = event.shiftKey ? buttons[0] : buttons.at(-1); if (document.activeElement === next) { event.preventDefault(); (event.shiftKey ? buttons.at(-1) : buttons[0])?.focus(); } }
    }}><h2 id="presentation-delete-title">Delete slide {index + 1}?</h2><p>You can undo this change before closing the presentation.</p><div><button autoFocus onClick={() => { setConfirmDelete(false); deleteTrigger.current?.focus(); }}>Cancel</button><button className="presentation-danger" onClick={() => { setConfirmDelete(false); structure('delete'); }}>Delete slide</button></div></section></div>}
  </div>;
}

/**
 * The painted run that carries the caret. React never owns its text: the run is
 * written once on mount and every keystroke is reported upwards, so typing never
 * moves the caret. Enter commits through the normal engine path, Esc cancels.
 */
function EditableRun({ text, style, offset, onInput, onCommit, onCancel, onBlur }: { text: string; style: CSSProperties; offset?: number; onInput(value: string): void; onCommit(): void; onCancel(): void; onBlur(keepEditing: boolean): void }) {
  const span = useRef<HTMLSpanElement>(null);
  useEffect(() => {
    const node = span.current; if (!node) return;
    node.textContent = text;
    node.focus({ preventScroll: true });
    // The click that started the edit settles its own selection first, so place the caret after it.
    const place = () => {
      const selection = window.getSelection(); if (!selection) return;
      const range = document.createRange();
      const at = Math.min(Math.max(offset ?? text.length, 0), text.length);
      if (node.firstChild) range.setStart(node.firstChild, at); else range.selectNodeContents(node);
      range.collapse(true);
      selection.removeAllRanges(); selection.addRange(range);
    };
    place();
    const timer = setTimeout(place, 0);
    return () => clearTimeout(timer);
    // eslint-disable-next-line react-hooks/exhaustive-deps -- the text is DOM state from here on.
  }, []);
  return <span ref={span} className="presentation-text-edit" role="textbox" aria-label="Slide text" contentEditable="plaintext-only" suppressContentEditableWarning spellCheck={false} style={style}
    onInput={event => onInput(event.currentTarget.textContent ?? '')}
    onKeyDown={event => {
      event.stopPropagation();
      if (event.key === 'Enter') { event.preventDefault(); onCommit(); }
      if (event.key === 'Escape') { event.preventDefault(); onCancel(); }
    }}
    onBlur={event => onBlur(!!(event.relatedTarget as Element | null)?.closest?.('.ribbon, .presentation-text-toolbar, .presentation-toolbar'))} />;
}

/**
 * One slide in the rail, laid out like PowerPoint: the number sits to the left of
 * the thumbnail and the stage takes the rest of the pane. The stage sizes itself
 * (`width: 100%` + the deck's `aspect-ratio`), so the thumbnail can never overflow
 * the pane; the painted slide is then scaled to the measured content box.
 */
function SlideThumbnail({ deck, slide, number, current, disabled, onSelect }: { deck: NativePptxDeck; slide: NativeSlide; number: number; current: boolean; disabled: boolean; onSelect(): void }) {
  const stage = useRef<HTMLDivElement>(null);
  const [width, setWidth] = useState(0);
  useEffect(() => {
    const node = stage.current; if (!node || typeof ResizeObserver !== 'function') return;
    const observer = new ResizeObserver(entries => setWidth(entries[0]?.contentRect.width ?? 0));
    observer.observe(node);
    return () => observer.disconnect();
  }, []);
  const slideWidth = deck.size.cx / EMU_PER_PIXEL;
  return <button className="presentation-thumbnail" disabled={disabled} aria-label={`Show slide ${number}`} aria-current={current ? 'page' : undefined} onClick={onSelect}>
    <span className="presentation-slide-number">{number}</span>
    <div className="presentation-thumb-stage" ref={stage} style={{ aspectRatio: `${deck.size.cx} / ${deck.size.cy}` }}>
      {width > 0 && slideWidth > 0 && <SlideCanvas deck={deck} slide={slide} scale={width / slideWidth} thumbnail />}
    </div>
  </button>;
}

function SlideCanvas({ deck, slide, scale, thumbnail = false, selected, selectedKeys = [], onSelect, disabled, draft, onGeometry, onGestureChange, selectedCell, onCellSelect, editing, onTextInput, onTextCommit, onTextCancel, onTextBlur }: { deck: NativePptxDeck; slide: NativeSlide; scale: number; thumbnail?: boolean; selected?: string; selectedKeys?: string[]; onSelect?(key: string, additive?: boolean, at?: { paragraph: number; run: number; offset?: number }): void; disabled?: boolean; draft?: Draft; onGeometry?(key: string, transform: NativeTransform): void; onGestureChange?(value: boolean): void; selectedCell?: { row: number; column: number }; onCellSelect?(key: string, row: number, column: number): void; editing?: { paragraph: number; run: number; offset?: number }; onTextInput?(value: string): void; onTextCommit?(): void; onTextCancel?(): void; onTextBlur?(keepEditing: boolean): void }) {
  const gesture = useRef<{ key: string; pointer: number; x: number; y: number; transform: NativeTransform; sx: number; sy: number; resize: boolean } | null>(null);
  const width = deck.size.cx / EMU_PER_PIXEL, height = deck.size.cy / EMU_PER_PIXEL;
  const elements = positionElements(slide.elements).map(item => {
    if (elementKey(item.element) !== selected || !draft) return item;
    if (draft.kind === 'text' && (item.element.kind === 'text' || item.element.kind === 'shape')) return { ...item, element: { ...item.element, paragraphs: draft.paragraphs.map(p => ({ ...p, runs: p.runs.map(r => ({ ...r })) })) } };
    if (draft.kind === 'table' && item.element.kind === 'table') {
      const rows = item.element.table.rows.map((row, r) => row.map((cell, c) => r === draft.row && c === draft.column ? { ...cell, text: draft.paragraphs.map(p => p.runs.map(run => run.text).join('')).join('\n'), paragraphs: draft.paragraphs.map(p => ({ ...p, runs: p.runs.map(r => ({ ...r })) })), fill: draft.fill } : cell));
      return { ...item, element: { ...item.element, table: { ...item.element.table, rows } } };
    }
    if (draft.kind === 'geometry' && (item.element.kind === 'text' || item.element.kind === 'picture' || item.element.kind === 'table')) { const t = draft.transform; return { ...item, element: { ...item.element, transform: t }, rect: { x: item.rect.x + (t.x - item.element.transform.x) * item.scaleX, y: item.rect.y + (t.y - item.element.transform.y) * item.scaleY, cx: t.cx * item.scaleX, cy: t.cy * item.scaleY } }; }
    if (draft.kind === 'rotation' && (item.element.kind === 'shape' || item.element.kind === 'picture' || item.element.kind === 'text')) return { ...item, element: { ...item.element, rotation60000: Number.isFinite(draft.rotation60000) ? draft.rotation60000 : item.element.rotation60000 } };
    if (draft.kind === 'shape' && item.element.kind === 'shape') { const shape = draft.shape; return { ...item, element: { ...item.element, ...shape }, rect: { x: item.rect.x + (shape.transform.x - item.element.transform.x) * item.scaleX, y: item.rect.y + (shape.transform.y - item.element.transform.y) * item.scaleY, cx: shape.transform.cx * item.scaleX, cy: shape.transform.cy * item.scaleY } }; }
    return item;
  });
  return <div className="presentation-slide-frame" style={{ width: width * scale, height: height * scale }}><div className="presentation-slide-surface" style={{ width, height, transform: `scale(${scale})`, background: slide.background ? `#${slide.background}` : '#ffffff' }} aria-label={thumbnail ? undefined : 'Slide canvas'}>
    {elements.map(({ element, rect, scaleX, scaleY }) => {
      const key = elementKey(element); const selectedHere = selectedKeys.length ? selectedKeys.includes(key) : selected === key;
      const movable = selectedKeys.length < 2 && !!onGeometry && (!!shapeTarget(deck, key) || !!geometryTarget(deck, key));
      const valid = [rect.x, rect.y, rect.cx, rect.cy].every(Number.isFinite) && rect.cx > 0 && rect.cy > 0;
      if (!valid) return null;
      const angle = shapeRotationDegrees(element);
      const style: CSSProperties = { transform: angle ? `rotate(${angle}deg)` : undefined, transformOrigin: 'center center', left: rect.x / EMU_PER_PIXEL, top: rect.y / EMU_PER_PIXEL, width: rect.cx / EMU_PER_PIXEL, height: rect.cy / EMU_PER_PIXEL };
      const textBody = (element.kind === 'text' || element.kind === 'shape') ? element.textBody : undefined;
      const fill = element.kind === 'shape' ? element.fill : undefined; const stroke = element.kind === 'shape' || element.kind === 'connector' ? element.stroke : undefined;
      const asset = element.kind === 'picture' ? deck.assets.find(value => value.id === element.assetId) : undefined;
      const paint = element.compatibility.status !== 'refused' && (element.kind !== 'picture' || element.compatibility.status === 'editable') && (element.kind !== 'shape' || ['rect', 'ellipse', 'triangle', 'diamond'].includes(element.preset ?? ''));
      return <div key={key} className={`presentation-object${selectedHere ? ' is-selected' : ''}${!paint ? ' is-placeholder' : ''}`} style={style} role={thumbnail ? undefined : element.kind === 'table' ? 'group' : 'button'} tabIndex={thumbnail || disabled ? -1 : 0} aria-label={thumbnail ? undefined : `Select ${element.name || element.kind}`} aria-pressed={thumbnail || element.kind === 'table' ? undefined : selectedHere} onPointerDown={event => {
          // A press on painted text puts the caret there instead of moving the object, so the pointer must not be captured.
          if (event.ctrlKey || event.metaKey || !selectedHere || disabled || !movable || event.button !== 0 || (event.target as Element).closest?.('[contenteditable], [data-slide-run]') || (element.kind === 'table' && (event.target as Element).closest('td'))) return;
          event.preventDefault(); event.stopPropagation(); event.currentTarget.setPointerCapture(event.pointerId); onGestureChange?.(true);
          gesture.current = { key, pointer: event.pointerId, x: event.clientX, y: event.clientY, transform: { ...element.transform }, sx: scale * scaleX / EMU_PER_PIXEL, sy: scale * scaleY / EMU_PER_PIXEL, resize: (event.target as HTMLElement).classList.contains('presentation-resize-handle') };
        }} onPointerMove={event => {
          const g = gesture.current; if (!g || event.pointerId !== g.pointer) return;
          const t = pointerTransform(g.transform, event.clientX - g.x, event.clientY - g.y, g.sx, g.sy, g.resize);
          onGeometry?.(g.key, t);
        }} onPointerUp={() => { gesture.current = null; onGestureChange?.(false); }} onPointerCancel={() => { gesture.current = null; onGestureChange?.(false); }} onLostPointerCapture={() => { gesture.current = null; onGestureChange?.(false); }}
        onClick={event => { event.stopPropagation(); if (disabled) return;
          const span = (event.target as Element).closest?.('[data-slide-run]') as HTMLElement | undefined;
          const paragraphs = (element.kind === 'text' || element.kind === 'shape') ? element.paragraphs : undefined;
          const last = paragraphs?.length ? { paragraph: paragraphs.length - 1, run: Math.max(0, (paragraphs.at(-1)?.runs.length ?? 1) - 1) } : undefined;
          const point = span ? (document as Document & { caretRangeFromPoint?(x: number, y: number): Range | null }).caretRangeFromPoint?.(event.clientX, event.clientY) : null;
          const offset = point && span?.firstChild === point.startContainer ? point.startOffset : undefined;
          const at = span ? { paragraph: Number(span.dataset.paragraph), run: Number(span.dataset.run), offset } : last;
          onSelect?.(key, event.ctrlKey || event.metaKey, at);
        }} onKeyDown={event => { if (!disabled && (event.key === 'Enter' || event.key === ' ')) { event.preventDefault(); onSelect?.(key, event.ctrlKey || event.metaKey); } }}>
        {selectedHere && movable && element.kind === 'table' && <span className="presentation-move-handle" aria-hidden="true">✥</span>}
        {selectedHere && movable && !angle && <span className="presentation-resize-handle" aria-hidden="true" />}
        {!paint ? <span className="presentation-object-placeholder">{element.name || element.kind} · preserved</span> : <>
          {(element.kind === 'shape' || element.kind === 'connector') && <svg className="presentation-shape-art" width="100%" height="100%" viewBox={`0 0 ${rect.cx / EMU_PER_PIXEL} ${rect.cy / EMU_PER_PIXEL}`} preserveAspectRatio="none"><ShapeArt element={element} width={rect.cx / EMU_PER_PIXEL} height={rect.cy / EMU_PER_PIXEL} fill={fill ? `#${fill}` : 'none'} stroke={stroke ? `#${stroke.color}` : 'none'} strokeWidth={stroke ? stroke.widthEmu * scaleX / EMU_PER_PIXEL : 0} /></svg>}
          {(element.kind === 'text' || element.kind === 'shape') && <div className="presentation-object-text" style={{ paddingLeft: (textBody?.leftInsetEmu ?? 0) * scaleX / EMU_PER_PIXEL, paddingRight: (textBody?.rightInsetEmu ?? 0) * scaleX / EMU_PER_PIXEL, paddingTop: (textBody?.topInsetEmu ?? 0) * scaleY / EMU_PER_PIXEL, paddingBottom: (textBody?.bottomInsetEmu ?? 0) * scaleY / EMU_PER_PIXEL, justifyContent: textBody?.verticalAnchor === 'center' ? 'center' : textBody?.verticalAnchor === 'bottom' ? 'flex-end' : 'flex-start', whiteSpace: textBody?.wrap === 'none' ? 'pre' : 'pre-wrap' }}>{element.paragraphs.map((p, pi) => <p key={pi} style={{ textAlign: p.align }}>{p.runs.map((r, ri) => {
            const runStyle: CSSProperties = { fontFamily: r.fontFamily, fontSize: (r.fontSizeHundredthPt ?? 1800) / 100 * 96 / 72 * scaleY, color: r.color ? `#${r.color}` : '#20242b', fontWeight: r.bold ? 'bold' : 'normal', fontStyle: r.italic ? 'italic' : 'normal' };
            if (!thumbnail && selectedHere && editing && editing.paragraph === pi && editing.run === ri) return <EditableRun key={ri} text={r.text} style={runStyle} offset={editing.offset} onInput={value => onTextInput?.(value)} onCommit={() => onTextCommit?.()} onCancel={() => onTextCancel?.()} onBlur={keepEditing => onTextBlur?.(keepEditing)} />;
            return <span key={ri} data-slide-run="" data-paragraph={pi} data-run={ri} style={runStyle}>{r.text || (element.paragraphs.length === 1 && p.runs.length === 1 && !thumbnail ? <span className="presentation-empty-text">Click to add text</span> : '')}</span>;
          })}</p>)}</div>}
          {element.kind === 'picture' && (asset?.dataBase64 && /^image\/(png|jpeg|gif|webp|bmp)$/.test(asset.contentType) ? <img draggable={false} src={`data:${asset.contentType};base64,${asset.dataBase64}`} alt={element.name || 'Slide image'} /> : <span className="presentation-object-placeholder">Image preserved</span>)}
          {element.kind === 'table' && <table className="presentation-table"><colgroup>{element.table.columnWidths.map((w, c) => <col key={c} style={{ width: `${w / element.table.columnWidths.reduce((a, b) => a + b, 0) * 100}%` }} />)}</colgroup><tbody>{element.table.rows.map((row, r) => <tr key={r} style={{ height: `${element.table.rowHeights[r]! / element.table.rowHeights.reduce((a, b) => a + b, 0) * 100}%` }}>{row.map((cell, c) => <td key={c} tabIndex={thumbnail || disabled ? -1 : 0} aria-label={thumbnail ? undefined : `Select table cell row ${r + 1} column ${c + 1}`} className={selectedHere && selectedCell?.row === r && selectedCell.column === c ? 'is-selected-cell' : undefined} onClick={event => { event.stopPropagation(); if (!disabled) onCellSelect?.(key, r, c); }} onKeyDown={event => { if (!disabled && (event.key === 'Enter' || event.key === ' ')) { event.preventDefault(); event.stopPropagation(); onCellSelect?.(key, r, c); } }} style={{ background: cell.fill ? `#${cell.fill}` : undefined }}><div className="presentation-cell-text" style={{ paddingLeft: (cell.textBody?.leftInsetEmu ?? 0) * scaleX / EMU_PER_PIXEL, paddingRight: (cell.textBody?.rightInsetEmu ?? 0) * scaleX / EMU_PER_PIXEL, paddingTop: (cell.textBody?.topInsetEmu ?? 0) * scaleY / EMU_PER_PIXEL, paddingBottom: (cell.textBody?.bottomInsetEmu ?? 0) * scaleY / EMU_PER_PIXEL }}>{cell.paragraphs ? cell.paragraphs.map((p, pi) => <p key={pi} style={{ textAlign: p.align }}>{p.runs.map((run, ri) => <span key={ri} style={{ fontFamily: run.fontFamily, fontSize: (run.fontSizeHundredthPt ?? 1800) / 100 * 96 / 72 * scaleY, fontWeight: run.bold ? 'bold' : 'normal', fontStyle: run.italic ? 'italic' : 'normal', color: run.color ? `#${run.color}` : undefined }}>{run.text}</span>)}</p>) : cell.text}</div></td>)}</tr>)}</tbody></table>}

          {(element.kind === 'chart' || element.kind === 'group') && <span className="presentation-object-placeholder">{element.kind} preserved</span>}
        </>}
      </div>;
    })}
  </div></div>;
}
