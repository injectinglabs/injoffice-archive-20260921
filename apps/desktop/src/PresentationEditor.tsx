import SlideArrangePanel from './SlideArrangePanel';
import { arrangeCommand, arrangeTargets, toggleArrangeSelection, type ArrangeAction } from './presentationArrange';
import ShapeArt from './ShapeArt';
import PresentationTextToolbar from './PresentationTextToolbar';
import { useEffect, useRef, useState, type CSSProperties } from 'react';
import { createPptxWasmClient, type PptxNativeExactAutoShapeV1, type PptxNativeExactParagraphV1, type PptxNativeMutationRequestV1 } from '@injoffice/pptx-wasm';
import type { NativeElement, NativePptxDeck, NativeSlide, NativeTransform } from '@injoffice/pptx-native';
import { EMU_PER_INCH, EMU_PER_PIXEL, elementKey, positionElements, shapeCommand, shapeTarget, structureCommand, textCommand, textTarget, recoveryDraft, restorePresentationDraft, insertCommand, deleteElementCommand, transformCommand, pointerTransform, geometryTarget, pictureInsertCommand, tableTarget, rotationCommand, rotationTarget, shapeRotationDegrees, slideBackgroundCommand, tableTopologyCommand, tableInsertCommand, tableCellCommand, tableTextParagraphs, tableSegmentParagraphs } from './presentationCommands';
import './presentation-editor.css';
import PresentationPlayer from './PresentationPlayer';
import { exportNativePptxSlideSvg } from '@injoffice/pptx-render';
import { startPresentationMode, type PresentationModeState } from './presentationMode';

export interface OfficeEditorProps { registerHistory?: (commands: { undo(): void; redo(): void }) => void; registerCommit?: (commit: () => Promise<boolean>) => void; initialRecoveryDraft?: unknown; onRecoveryDraftChange?: (draft: unknown | null) => void; name: string; bytes: Uint8Array; onChange: (bytes: Uint8Array) => void; onBusyChange?: (busy: boolean) => void; onDraftChange?: (dirty: boolean) => void; viewOptions?: { zoom: number; navigation: boolean; focus: boolean } }
type Snapshot = { bytes: Uint8Array; deck: NativePptxDeck };
type Draft = import('./presentationCommands').PresentationDraft;
type PresentationEditorProps = OfficeEditorProps & { initialRecoveryDraft?: unknown; onRecoveryDraftChange?(draft: unknown | null): void; registerCommit?(commit: () => Promise<boolean>): void; registerHistory?(commands: { undo(): void; redo(): void }): void };
const newId = () => `slides-${crypto.randomUUID()}`;
const describeError = (error: unknown) => error instanceof Error ? error.message : String(error);

export default function PresentationEditor({ name, bytes, onChange, onBusyChange, onDraftChange, viewOptions, initialRecoveryDraft, onRecoveryDraftChange, registerCommit, registerHistory }: PresentationEditorProps) {
  const [presentation, setPresentation] = useState<PresentationModeState>();
  const presenting = useRef(false);
  const historyRef = useRef<(direction: 'undo' | 'redo') => void>(() => {});
  useEffect(() => { registerHistory?.({ undo: () => historyRef.current('undo'), redo: () => historyRef.current('redo') }); }, [registerHistory]);
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
  const [panel, setPanel] = useState<'text' | 'shape' | 'position' | 'table' | 'rotation'>('text');
  const [cellSelection, setCellSelection] = useState({ row: 0, column: 0 });
  const [cellSegment, setCellSegment] = useState({ paragraph: 0, run: 0 });
  const [cellTextEdit, setCellTextEdit] = useState<{ text: string; paragraphs: readonly PptxNativeExactParagraphV1[] }>();
  const [tableSize, setTableSize] = useState({ rows: 3, columns: 3 });
  const [tableInsertOpen, setTableInsertOpen] = useState(false);
  const [segment, setSegment] = useState({ paragraph: 0, run: 0 });
  const [error, setError] = useState('');
  const [confirmDelete, setConfirmDelete] = useState(false);
  const deleteTrigger = useRef<HTMLButtonElement>(null);
  const [undo, setUndo] = useState<Snapshot[]>([]); const [redo, setRedo] = useState<Snapshot[]>([]);
  const undoRef = useRef<Snapshot[]>([]); const redoRef = useRef<Snapshot[]>([]);
  const mounted = useRef(false);
  const workspace = useRef<HTMLDivElement>(null); const [workspaceWidth, setWorkspaceWidth] = useState(740);
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
        setIndex(target!.slideIndex); setSelected(recovered.key); setPanel(recovered.draft.kind === 'geometry' ? 'position' : recovered.draft.kind);
        draftRef.current = recovered.draft; setDraft(recovered.draft); callbacks.current.onDraftChange?.(true); callbacks.current.onBusyChange?.(busyRef.current);
      } } catch (reason) { setError(describeError(reason)); }
    } })
      .catch(reason => { if (!cancelled) setError(describeError(reason)); })
      .finally(() => { if (!cancelled) markBusy(false); });
    return () => { cancelled = true; mounted.current = false; worker.terminate(); callbacks.current.onBusyChange?.(false); callbacks.current.onDraftChange?.(false); };
  }, []);
  useEffect(() => {
    if (!workspace.current) return;
    const observer = new ResizeObserver(entries => setWorkspaceWidth(entries[0]?.contentRect.width ?? 740)); observer.observe(workspace.current);
    return () => observer.disconnect();
  }, [!!snapshot, viewOptions?.focus]);
  const blocked = busy || !!draft || confirmDelete;
  const slide = snapshot?.deck.slides[index];
  const selectedItem = slide && positionElements(slide.elements).find(item => elementKey(item.element) === selected);
  const text = snapshot && textTarget(snapshot.deck, selected);
  const shape = snapshot && shapeTarget(snapshot.deck, selected);
  const table = snapshot && tableTarget(snapshot.deck, selected);
  const tableCell = table?.element.table.rows[cellSelection.row]?.[cellSelection.column];
  const cellDraft = draft?.kind === 'table' ? draft : undefined;
  const paragraphs = draft?.kind === 'text' ? draft.paragraphs : text?.paragraphs;
  const selectedParagraph = paragraphs?.[segment.paragraph]; const run = selectedParagraph?.runs[segment.run];
  const shapeValue = draft?.kind === 'shape' ? draft.shape : shape?.autoShape;
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
    setSelected(key); setCellSelection({ row: 0, column: 0 }); setSegment({ paragraph: 0, run: 0 }); setPanel(current.current && tableTarget(current.current.deck, key) ? 'table' : current.current && textTarget(current.current.deck, key) ? 'text' : current.current && geometryTarget(current.current.deck, key)?.element.kind === 'picture' ? 'position' : 'shape'); setError('');
  }
  function arrange(action: ArrangeAction) {
    if (blocked || !current.current || dragging.current || composing.current) return;
    try { void applyCommand(arrangeCommand(current.current.deck,index,arrangeKeys,action,newId())); }
    catch (reason) { setError(describeError(reason)); }
  }
  function selectSlide(next: number) { if (blocked) return; setIndex(next); setSelected(''); setArrangeKeys([]); setSegment({ paragraph: 0, run: 0 }); setError(''); }
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
        const item = deck.slides[nextIndex]?.elements.at(-1); if (item) { setArrangeKeys([]); setSelected(elementKey(item)); setPanel(request.operations[0].kind === 'table.insert' ? 'table' : request.operations[0].kind === 'text.insert' ? 'text' : request.operations[0].kind === 'picture.insert' ? 'position' : 'shape'); setSegment({ paragraph: 0, run: 0 }); setCellSelection({ row: 0, column: 0 }); }
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
  const scale = snapshot ? Math.max(.1, Math.min(1, (workspaceWidth - 64) / (snapshot.deck.size.cx / EMU_PER_PIXEL))) * Math.max(.5, Math.min(2, (viewOptions?.zoom ?? 100) / 100)) : 1;
  return <div className="presentation-editor" aria-label="Presentation editor" aria-busy={busy} onKeyDown={event => {
      if (!(event.ctrlKey || event.metaKey) || event.altKey || blocked || composing.current || dragging.current) return;
      const key = event.key.toLowerCase(); if (key !== 'z' && key !== 'y') return;
      event.preventDefault(); event.stopPropagation(); travel(key === 'y' || event.shiftKey ? 'redo' : 'undo');
    }} onCompositionStart={() => { composing.current = true; }} onCompositionEnd={() => { composing.current = false; }}>
    {presentation && <PresentationPlayer initial={presentation} onExit={() => { presenting.current = false; setPresentation(undefined); }} renderSlide={(slide, scale) => <SlideCanvas deck={presentation.deck} slide={slide} scale={scale} thumbnail />} />}
    <div className="presentation-toolbar">
      <button disabled={busy || !snapshot || confirmDelete} onClick={() => { if (!current.current || busyRef.current || dragging.current || composing.current) return; presenting.current = true; setPresentation(startPresentationMode(current.current.deck, index, !!draftRef.current)); }}>Present</button>
      <div><button disabled={blocked || !undo.length} onClick={() => travel('undo')}>Undo</button><button disabled={blocked || !redo.length} onClick={() => travel('redo')}>Redo</button></div>
      <div><button disabled={blocked || !slide} onClick={() => insert('slide')}>New slide</button><button disabled={blocked || !slide} onClick={() => structure('duplicate')}>Duplicate slide</button><button disabled={blocked || !slide || index === 0} onClick={() => structure('previous')}>Move earlier</button><button disabled={blocked || !snapshot || index >= snapshot.deck.slides.length - 1} onClick={() => structure('next')}>Move later</button><button disabled={blocked || !snapshot || snapshot.deck.slides.length <= 1} ref={deleteTrigger} onClick={() => setConfirmDelete(true)}>Delete slide</button></div>
      <div><button disabled={blocked || !slide} onClick={() => insert('text')}>Text box</button><button disabled={blocked || !slide} onClick={() => insert('shape')}>Shape</button><button disabled={blocked || !slide} onClick={() => void insertPicture()}>Picture</button><button disabled={blocked || !slide} aria-expanded={tableInsertOpen} onClick={() => setTableInsertOpen(!tableInsertOpen)}>Table</button><button disabled={blocked || !selectedItem || selectedItem.grouped || (!text && !shape && !geometryTarget(snapshot!.deck, selected))} onClick={deleteObject}>Delete object</button></div>
      <button disabled={blocked || !snapshot} onClick={() => void exportSvg()} title="Export the current supported slide. Text and unsupported objects require the original PPTX.">Export slide SVG</button>
      <label className="presentation-background-control">Background<select aria-label="Slide background" value={slide?.background ?? ''} disabled={blocked || !slide || slide.compatibility.diagnostics.some(d => d.code === 'pptx.unsupported-background')} onChange={event => changeBackground(event.target.value)}>
        <option value="" disabled>Inherited</option>
        {slide?.background && !['FFFFFF','F5F7FA','202B3C','2459AD','DCE8F7','E4F1E9','FFF2D2','F6E3E6'].includes(slide.background) && <option value={slide.background}>Custom #{slide.background}</option>}
        {Object.entries({FFFFFF:'White',F5F7FA:'Fog', '202B3C':'Ink','2459AD':'Blue',DCE8F7:'Pale blue',E4F1E9:'Sage',FFF2D2:'Cream',F6E3E6:'Rose'}).map(([color,label]) => <option key={color} value={color}>{label}</option>)}
      </select></label>
      {tableInsertOpen && <div className="presentation-table-options" aria-label="Insert table"><label>Rows <input type="number" min="1" max="100" aria-label="New table rows" value={Number.isFinite(tableSize.rows) ? tableSize.rows : ''} onChange={event => setTableSize({ ...tableSize, rows: event.target.valueAsNumber })} /></label><label>Columns <input type="number" min="1" max="100" aria-label="New table columns" value={Number.isFinite(tableSize.columns) ? tableSize.columns : ''} onChange={event => setTableSize({ ...tableSize, columns: event.target.valueAsNumber })} /></label><button disabled={blocked} onClick={insertTable}>Insert table</button><button onClick={() => setTableInsertOpen(false)}>Cancel</button></div>}
      {draft && <div className="presentation-pending"><span>Pending changes</span><button className="presentation-primary" disabled={busy} onClick={applyDraft}>Apply changes</button><button disabled={busy} onClick={() => { updateDraft(undefined); setError(''); }}>Cancel</button></div>}
    </div>
    <PresentationTextToolbar run={run} align={selectedParagraph?.align} disabled={busy || confirmDelete || !text || (!!draft && draft.kind !== 'text')} onRunChange={patch => textPatch(patch)} onAlignChange={align => textPatch({}, align)} />
    {error && <div className="presentation-error" role="alert"><span>{error}</span><button aria-label="Dismiss presentation error" onClick={() => setError('')}>×</button></div>}
    {!snapshot ? <div className="presentation-loading" role="status">{busy ? 'Opening presentation…' : 'This presentation could not be opened.'}</div> : <div className="presentation-layout">
      {!viewOptions?.focus && <nav className="presentation-thumbnails" aria-label="Slides"><div className="presentation-rail-title">Slides <span>{snapshot.deck.slides.length}</span></div>{snapshot.deck.slides.map((item, i) => <button key={item.id} className="presentation-thumbnail" disabled={blocked} aria-label={`Show slide ${i + 1}`} aria-current={index === i ? 'page' : undefined} onClick={() => selectSlide(i)}><span className="presentation-slide-number">{i + 1}</span><div className="presentation-thumb-stage"><SlideCanvas deck={snapshot.deck} slide={item} scale={148 / (snapshot.deck.size.cx / EMU_PER_PIXEL)} thumbnail /></div></button>)}</nav>}
      <div className="presentation-workspace" ref={workspace}>
        <div className="presentation-canvas-label"><strong>Slide {index + 1}</strong><span>Positioned preview · text wrapping may differ in PowerPoint</span></div>
        <div className="presentation-canvas-scroll">{slide && <SlideCanvas deck={snapshot.deck} slide={slide} scale={scale} selected={selected} selectedKeys={arrangeKeys} onSelect={choose} disabled={blocked} draft={draft} onGeometry={geometryPatch} selectedCell={cellSelection} onCellSelect={chooseCell} onGestureChange={value => { dragging.current = value; }} />}</div>
        {slide && <details className="presentation-fidelity"><summary>Preview and editing limits</summary><p>Exact text and supported shapes can be edited. Images use embedded previews when available. Unsupported content stays in the file and appears as a placeholder. Slide commands can be refused for notes, comments, links, sections, or other relationships that cannot be changed safely.</p>{slide.compatibility.diagnostics.length > 0 && <ul>{slide.compatibility.diagnostics.slice(0, 10).map((diagnostic, i) => <li key={i}>{diagnostic.message}</li>)}</ul>}</details>}
      </div>
      <aside className="presentation-inspector" aria-label="Selected object">
        <SlideArrangePanel elements={arrangeTargets(snapshot.deck,index)} keys={arrangeKeys} disabled={blocked} onToggle={key => choose(key,true)} onArrange={arrange} />
        {arrangeKeys.length > 1 ? <p className="presentation-help">Choose a single object to edit its content or appearance.</p> : <>
        <h2>{selectedItem?.element.name || (selectedItem ? selectedItem.element.kind : 'Select an object')}</h2>
        {!selectedItem ? <p className="presentation-help">Choose text or a shape on the slide to edit it.</p> : <>
          <div className="presentation-inspector-tabs">{rotationTarget(snapshot.deck, selected) && <button aria-pressed={panel === 'rotation'} disabled={!!draft || busy} onClick={() => setPanel('rotation')}>Rotation</button>}{table && <button aria-pressed={panel === 'table'} disabled={!!draft || busy} onClick={() => setPanel('table')}>Cell</button>}{snapshot && geometryTarget(snapshot.deck, selected) && <button aria-pressed={panel === 'position'} disabled={!!draft || busy} onClick={() => setPanel('position')}>Position</button>}{text && <button aria-pressed={panel === 'text'} disabled={!!draft || busy} onClick={() => setPanel('text')}>Text</button>}{shape && <button aria-pressed={panel === 'shape'} disabled={!!draft || busy} onClick={() => setPanel('shape')}>Shape</button>}</div>
          {!text && !shape && !geometryTarget(snapshot.deck, selected) && <p className="presentation-help">This object is preserved in the original file. Editing is not available for its current format.</p>}
          {panel === 'text' && text && paragraphs && <fieldset disabled={busy}>
            <label>Text segment<select aria-label="Slide text segment" value={`${segment.paragraph}:${segment.run}`} onChange={event => { const [paragraph, run] = event.target.value.split(':').map(Number); setSegment({ paragraph: paragraph!, run: run! }); }}>{paragraphs.flatMap((p, pi) => p.runs.map((r, ri) => <option key={`${pi}:${ri}`} value={`${pi}:${ri}`}>Paragraph {pi + 1}, segment {ri + 1}: {r.text.slice(0, 35) || 'Empty text'}</option>))}</select></label>
            {run && <><label>Text<textarea aria-label="Slide text" rows={5} value={run.text} onChange={event => textPatch({ text: event.target.value })} /></label>
              <label>Font<input aria-label="Slide font family" list="presentation-fonts" value={run.fontFamily} onChange={event => textPatch({ fontFamily: event.target.value })} /><datalist id="presentation-fonts">{['Arial', 'Calibri', 'Georgia', 'Times New Roman', 'Verdana'].map(font => <option key={font}>{font}</option>)}</datalist></label>
              <div className="presentation-inspector-row"><label>Size (pt)<input type="number" min="1" max="400" step=".5" aria-label="Slide font size" value={Number.isFinite(run.fontSizeHundredthPt) ? run.fontSizeHundredthPt / 100 : ''} onChange={event => textPatch({ fontSizeHundredthPt: Math.round(event.target.valueAsNumber * 100) })} /></label><label>Color<input type="color" aria-label="Slide text color" value={`#${run.color}`} onChange={event => textPatch({ color: event.target.value.slice(1).toUpperCase() })} /></label></div>
              <div className="presentation-text-buttons"><button aria-label="Slide bold" aria-pressed={run.bold} onClick={() => textPatch({ bold: !run.bold })}><b>B</b></button><button aria-label="Slide italic" aria-pressed={run.italic} onClick={() => textPatch({ italic: !run.italic })}><i>I</i></button><select aria-label="Slide paragraph alignment" value={selectedParagraph?.align} onChange={event => textPatch({}, event.target.value as 'left' | 'center' | 'right')}><option value="left">Align left</option><option value="center">Center</option><option value="right">Align right</option></select></div>
              <p className="presentation-help">Text formatting changes this segment. Alignment changes its paragraph. New line breaks are not supported inside a segment.</p>
            </>}
          </fieldset>}
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
      </aside>
    </div>}
    {confirmDelete && <div className="presentation-modal"><section role="alertdialog" aria-modal="true" aria-labelledby="presentation-delete-title" onKeyDown={event => {
      if (event.key === 'Escape') { event.preventDefault(); setConfirmDelete(false); deleteTrigger.current?.focus(); }
      if (event.key === 'Tab') { const buttons = Array.from(event.currentTarget.querySelectorAll('button')); const next = event.shiftKey ? buttons[0] : buttons.at(-1); if (document.activeElement === next) { event.preventDefault(); (event.shiftKey ? buttons.at(-1) : buttons[0])?.focus(); } }
    }}><h2 id="presentation-delete-title">Delete slide {index + 1}?</h2><p>You can undo this change before closing the presentation.</p><div><button autoFocus onClick={() => { setConfirmDelete(false); deleteTrigger.current?.focus(); }}>Cancel</button><button className="presentation-danger" onClick={() => { setConfirmDelete(false); structure('delete'); }}>Delete slide</button></div></section></div>}
  </div>;
}

function SlideCanvas({ deck, slide, scale, thumbnail = false, selected, selectedKeys = [], onSelect, disabled, draft, onGeometry, onGestureChange, selectedCell, onCellSelect }: { deck: NativePptxDeck; slide: NativeSlide; scale: number; thumbnail?: boolean; selected?: string; selectedKeys?: string[]; onSelect?(key: string, additive?: boolean): void; disabled?: boolean; draft?: Draft; onGeometry?(key: string, transform: NativeTransform): void; onGestureChange?(value: boolean): void; selectedCell?: { row: number; column: number }; onCellSelect?(key: string, row: number, column: number): void }) {
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
          if (event.ctrlKey || event.metaKey || !selectedHere || disabled || !movable || event.button !== 0 || (element.kind === 'table' && (event.target as Element).closest('td'))) return;
          event.preventDefault(); event.stopPropagation(); event.currentTarget.setPointerCapture(event.pointerId); onGestureChange?.(true);
          gesture.current = { key, pointer: event.pointerId, x: event.clientX, y: event.clientY, transform: { ...element.transform }, sx: scale * scaleX / EMU_PER_PIXEL, sy: scale * scaleY / EMU_PER_PIXEL, resize: (event.target as HTMLElement).classList.contains('presentation-resize-handle') };
        }} onPointerMove={event => {
          const g = gesture.current; if (!g || event.pointerId !== g.pointer) return;
          const t = pointerTransform(g.transform, event.clientX - g.x, event.clientY - g.y, g.sx, g.sy, g.resize);
          onGeometry?.(g.key, t);
        }} onPointerUp={() => { gesture.current = null; onGestureChange?.(false); }} onPointerCancel={() => { gesture.current = null; onGestureChange?.(false); }} onLostPointerCapture={() => { gesture.current = null; onGestureChange?.(false); }}
        onClick={event => { event.stopPropagation(); if (!disabled) onSelect?.(key, event.ctrlKey || event.metaKey); }} onKeyDown={event => { if (!disabled && (event.key === 'Enter' || event.key === ' ')) { event.preventDefault(); onSelect?.(key, event.ctrlKey || event.metaKey); } }}>
        {selectedHere && movable && element.kind === 'table' && <span className="presentation-move-handle" aria-hidden="true">✥</span>}
        {selectedHere && movable && !angle && <span className="presentation-resize-handle" aria-hidden="true" />}
        {!paint ? <span className="presentation-object-placeholder">{element.name || element.kind} · preserved</span> : <>
          {(element.kind === 'shape' || element.kind === 'connector') && <svg className="presentation-shape-art" width="100%" height="100%" viewBox={`0 0 ${rect.cx / EMU_PER_PIXEL} ${rect.cy / EMU_PER_PIXEL}`} preserveAspectRatio="none"><ShapeArt element={element} width={rect.cx / EMU_PER_PIXEL} height={rect.cy / EMU_PER_PIXEL} fill={fill ? `#${fill}` : 'none'} stroke={stroke ? `#${stroke.color}` : 'none'} strokeWidth={stroke ? stroke.widthEmu * scaleX / EMU_PER_PIXEL : 0} /></svg>}
          {(element.kind === 'text' || element.kind === 'shape') && <div className="presentation-object-text" style={{ paddingLeft: (textBody?.leftInsetEmu ?? 0) * scaleX / EMU_PER_PIXEL, paddingRight: (textBody?.rightInsetEmu ?? 0) * scaleX / EMU_PER_PIXEL, paddingTop: (textBody?.topInsetEmu ?? 0) * scaleY / EMU_PER_PIXEL, paddingBottom: (textBody?.bottomInsetEmu ?? 0) * scaleY / EMU_PER_PIXEL, justifyContent: textBody?.verticalAnchor === 'center' ? 'center' : textBody?.verticalAnchor === 'bottom' ? 'flex-end' : 'flex-start', whiteSpace: textBody?.wrap === 'none' ? 'pre' : 'pre-wrap' }}>{element.paragraphs.map((p, pi) => <p key={pi} style={{ textAlign: p.align }}>{p.runs.map((r, ri) => <span key={ri} style={{ fontFamily: r.fontFamily, fontSize: (r.fontSizeHundredthPt ?? 1800) / 100 * 96 / 72 * scaleY, color: r.color ? `#${r.color}` : '#20242b', fontWeight: r.bold ? 'bold' : 'normal', fontStyle: r.italic ? 'italic' : 'normal' }}>{r.text || (element.paragraphs.length === 1 && p.runs.length === 1 && !thumbnail ? <span className="presentation-empty-text">Select to add text</span> : '')}</span>)}</p>)}</div>}
          {element.kind === 'picture' && (asset?.dataBase64 && /^image\/(png|jpeg|gif|webp|bmp)$/.test(asset.contentType) ? <img draggable={false} src={`data:${asset.contentType};base64,${asset.dataBase64}`} alt={element.name || 'Slide image'} /> : <span className="presentation-object-placeholder">Image preserved</span>)}
          {element.kind === 'table' && <table className="presentation-table"><colgroup>{element.table.columnWidths.map((w, c) => <col key={c} style={{ width: `${w / element.table.columnWidths.reduce((a, b) => a + b, 0) * 100}%` }} />)}</colgroup><tbody>{element.table.rows.map((row, r) => <tr key={r} style={{ height: `${element.table.rowHeights[r]! / element.table.rowHeights.reduce((a, b) => a + b, 0) * 100}%` }}>{row.map((cell, c) => <td key={c} tabIndex={thumbnail || disabled ? -1 : 0} aria-label={thumbnail ? undefined : `Select table cell row ${r + 1} column ${c + 1}`} className={selectedHere && selectedCell?.row === r && selectedCell.column === c ? 'is-selected-cell' : undefined} onClick={event => { event.stopPropagation(); if (!disabled) onCellSelect?.(key, r, c); }} onKeyDown={event => { if (!disabled && (event.key === 'Enter' || event.key === ' ')) { event.preventDefault(); event.stopPropagation(); onCellSelect?.(key, r, c); } }} style={{ background: cell.fill ? `#${cell.fill}` : undefined }}><div className="presentation-cell-text" style={{ paddingLeft: (cell.textBody?.leftInsetEmu ?? 0) * scaleX / EMU_PER_PIXEL, paddingRight: (cell.textBody?.rightInsetEmu ?? 0) * scaleX / EMU_PER_PIXEL, paddingTop: (cell.textBody?.topInsetEmu ?? 0) * scaleY / EMU_PER_PIXEL, paddingBottom: (cell.textBody?.bottomInsetEmu ?? 0) * scaleY / EMU_PER_PIXEL }}>{cell.paragraphs ? cell.paragraphs.map((p, pi) => <p key={pi} style={{ textAlign: p.align }}>{p.runs.map((run, ri) => <span key={ri} style={{ fontFamily: run.fontFamily, fontSize: (run.fontSizeHundredthPt ?? 1800) / 100 * 96 / 72 * scaleY, fontWeight: run.bold ? 'bold' : 'normal', fontStyle: run.italic ? 'italic' : 'normal', color: run.color ? `#${run.color}` : undefined }}>{run.text}</span>)}</p>) : cell.text}</div></td>)}</tr>)}</tbody></table>}

          {(element.kind === 'chart' || element.kind === 'group') && <span className="presentation-object-placeholder">{element.kind} preserved</span>}
        </>}
      </div>;
    })}
  </div></div>;
}
