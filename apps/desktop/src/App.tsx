import { lazy, Suspense, useCallback, useEffect, useRef, useState } from 'react';
import { flushSync } from 'react-dom';
import OfficeEditor from './OfficeEditor';
const PdfEditor = lazy(() => import('./PdfEditor'));
const PresentationEditor = lazy(() => import('./PresentationEditor'));
const SpreadsheetEditor = lazy(() => import('./SpreadsheetEditor').then(module => ({ default: module.SpreadsheetEditor })));
import StartPage from './StartPage';
import OpenError, { classifyOpenError, containerFailure, type OpenErrorKind } from './OpenError';
import UpdatesDialog, { UpdateNotice } from './UpdatesDialog';
import PreferencesDialog from './PreferencesDialog';
import { readPreferences, writePreferences, initialView, type ViewOptions } from './preferences';
import { applyTheme } from './theme';
import CommandPalette, { type WorkspaceCommand } from './CommandPalette';
import { RibbonButton, RibbonRows, WorkspaceFileGroupsContext, type WorkspaceFileGroups } from './Ribbon';
import RibbonIcon from './RibbonIcons';
import { shortcutLabel, shortcutPlatform, shortcutTooltip } from './shortcuts';

// macOS hides the system title bar (main.cjs uses titleBarStyle 'hiddenInset'), so the app title bar
// is the window's only title row: it drags the window and leaves room for the inset traffic lights.
const macChrome = shortcutPlatform() === 'mac';

type DocumentSession = { key: number; id: string; name: string; initialName: string; initialBytes: Uint8Array; bytes: Uint8Array; dirty: boolean; untitled?: boolean; editorBusy?: boolean; draftDirty?: boolean; recoveryDraft?: unknown };
type ReplaceChoice = 'save' | 'discard' | 'cancel';

type HistoryCommands = { undo(): void; redo(): void; canUndo?: boolean; canRedo?: boolean };

function SessionEditor({ session, onSessionChange, viewOptions, registerSessionCommit, registerSessionHistory }: { registerSessionHistory(key: number, commands: HistoryCommands): void; session: DocumentSession; registerSessionCommit(key: number, commit: () => Promise<boolean>): void; onSessionChange(key: number, patch: Partial<DocumentSession>): void; viewOptions: { zoom: number; navigation: boolean; focus: boolean } }) {
  const key = session.key;
  const onChange = useCallback((bytes: Uint8Array) => onSessionChange(key, { bytes: new Uint8Array(bytes), dirty: true, recoveryDraft: null }), [key, onSessionChange]);
  const onBusyChange = useCallback((editorBusy: boolean) => onSessionChange(key, { editorBusy }), [key, onSessionChange]);
  const onDraftChange = useCallback((draftDirty: boolean) => onSessionChange(key, { draftDirty }), [key, onSessionChange]);
  const registerHistory = useCallback((commands: HistoryCommands) => registerSessionHistory(key, commands), [key, registerSessionHistory]);
  const registerCommit = useCallback((commit: () => Promise<boolean>) => registerSessionCommit(key, commit), [key, registerSessionCommit]);
  const onRecoveryDraftChange = useCallback((recoveryDraft: unknown | null) => onSessionChange(key, { recoveryDraft }), [key, onSessionChange]);
  const format = session.initialName.split('.').pop()?.toLowerCase();
  const Editor = format === 'pdf' ? PdfEditor : format === 'pptx' ? PresentationEditor : format === 'xlsx' ? SpreadsheetEditor : OfficeEditor;
  return <Suspense fallback={<div className="office-empty" role="status">Opening editor…</div>}><Editor registerHistory={registerHistory} registerCommit={registerCommit} initialRecoveryDraft={session.recoveryDraft} onRecoveryDraftChange={onRecoveryDraftChange} name={session.initialName} bytes={session.initialBytes} onChange={onChange} onBusyChange={onBusyChange} onDraftChange={onDraftChange} viewOptions={viewOptions} /></Suspense>;
}

export default function App() {
  const [document, setDocument] = useState<DocumentSession | null>(null);
  const [sessions, setSessions] = useState<DocumentSession[]>([]);
  const historyHandlers = useRef(new Map<number, HistoryCommands>());
  // Editors that report canUndo/canRedo grey the title bar's Undo/Redo out; the rest keep them enabled like the Edit menu.
  const [historyState, setHistoryState] = useState<Record<number, { undo: boolean; redo: boolean }>>({});
  const registerSessionHistory = useCallback((key: number, commands: HistoryCommands) => {
    historyHandlers.current.set(key, commands);
    if (commands.canUndo === undefined && commands.canRedo === undefined) return;
    const undo = commands.canUndo ?? true, redo = commands.canRedo ?? true;
    setHistoryState(previous => previous[key]?.undo === undo && previous[key]?.redo === redo ? previous : { ...previous, [key]: { undo, redo } });
  }, []);
  const commitHandlers = useRef(new Map<number, () => Promise<boolean>>());
  const registerSessionCommit = useCallback((key: number, commit: () => Promise<boolean>) => { commitHandlers.current.set(key, commit); }, []);
  const sessionsRef = useRef<DocumentSession[]>([]);
  const [recoveries, setRecoveries] = useState<Array<{ id: string; name: string; updatedAt: number; size: number }>>([]);
  const current = useRef<DocumentSession | null>(null);
  const sequence = useRef(0);
  const checkpointSources = useRef(new Map<string,{bytes:Uint8Array;revision:string}>());
  const pendingCheckpoints = useRef(new Set<Promise<void>>());
  const closingRef = useRef(false);
  const rootElement = useRef<HTMLDivElement>(null);
  const [closing, setClosing] = useState(false);
  const externalPending = useRef(true);
  const operation = useRef(false);
  const editorBusyRef = useRef(false);
  const draftDirtyRef = useRef(false);
  const [draftDirty, setDraftDirty] = useState(false);
  const [working, setWorking] = useState(false);
  const [importProgress,setImportProgress]=useState<{completed:number;total:number}|null>(null);
  const importAbort=useRef<AbortController|undefined>(undefined);
  useEffect(()=>()=>importAbort.current?.abort(),[]);
  const [editorBusy, setEditorBusy] = useState(false);
  const [error, setError] = useState('');
  const [openFailure, setOpenFailure] = useState<{ kind: OpenErrorKind; detail: string; name?: string } | null>(null);
  const [notice, setNotice] = useState('');
  const [replacePrompt, setReplacePrompt] = useState(false);
  const [showHome, setShowHome] = useState(true);
  const [commandSearch, setCommandSearch] = useState(false);
  const [preferences, setPreferences] = useState(readPreferences);
  const preferencesRef = useRef(preferences); preferencesRef.current = preferences;
  const [updatesOpen, setUpdatesOpen] = useState(false);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const sessionViews = useRef(new Map<number, ViewOptions>());
  const [recentFiles, setRecentFiles] = useState<Array<{ id: string; name: string; path: string; updatedAt: number }>>([]);
  const [viewOptions, publishViewOptions] = useState(() => initialView(preferences));
  const setViewOptions = (update: (previous: ViewOptions) => ViewOptions) => {
    publishViewOptions(previous => { const next = update(previous); if (current.current) sessionViews.current.set(current.current.key, next); return next; });
  };
  const resolveChoice = useRef<((choice: ReplaceChoice) => void) | null>(null);
  const bridge = window.injDesktop;
  useEffect(() => { applyTheme(preferences.theme); void bridge?.setTheme?.(preferences.theme).catch(() => { /* Native dialogs keep the previous appearance; the renderer already switched. */ }); }, [bridge, preferences.theme]);

  const refreshRecent = useCallback(async () => {
    if (!bridge) return;
    try { setRecentFiles(await bridge.recent()); }
    catch { setError('Recent files could not be loaded. You can still use Open file.'); }
  }, [bridge]);

  useEffect(() => { void refreshRecent(); void bridge?.recovery().then(setRecoveries).catch(() => setError('Recovery copies could not be loaded.')); }, [refreshRecent, bridge]);

  const removeRecent = async (id: string) => {
    try { await bridge?.removeRecent(id); await refreshRecent(); }
    catch { setError('This file could not be removed from recent files. Try again.'); }
  };

  const publishSessions = useCallback((next: DocumentSession[], activeKey = current.current?.key) => {
    sessionsRef.current = next; setSessions(next);
    const active = next.find(item => item.key === activeKey) ?? next.at(-1) ?? null;
    current.current = active; setDocument(active);
    if (active) {
      const view = sessionViews.current.get(active.key) ?? initialView(preferencesRef.current);
      sessionViews.current.set(active.key, view); publishViewOptions(view);
    }
    editorBusyRef.current = active?.editorBusy ?? false; draftDirtyRef.current = active?.draftDirty ?? false;
    setEditorBusy(editorBusyRef.current); setDraftDirty(draftDirtyRef.current);
    window.injDesktop?.setDirty(next.some(item => item.dirty || item.draftDirty));
    window.injDesktop?.setBusy?.(next.some(item => item.editorBusy));
  }, []);
  const updateDocument = useCallback((next: DocumentSession | null) => {
    if (!next) { publishSessions([]); return; }
    const existing = sessionsRef.current;
    publishSessions(existing.some(item => item.key === next.key) ? existing.map(item => item.key === next.key ? next : item) : [...existing, next], next.key);
  }, [publishSessions]);
  const checkpoint = useCallback((session: DocumentSession) => {
    const previous=checkpointSources.current.get(session.id),same=previous?.bytes===session.bytes;
    const source=same?previous!:{bytes:session.bytes,revision:crypto.randomUUID()};checkpointSources.current.set(session.id,source);
    const pending = window.injDesktop?.checkpoint(session.id,same?null:session.bytes,session.recoveryDraft??null,source.revision).catch(cause=>{
      if(checkpointSources.current.get(session.id)===source)checkpointSources.current.delete(session.id);
      setError(`Recovery could not be updated: ${cause instanceof Error ? cause.message : String(cause)}`);
    });
    if (pending) { pendingCheckpoints.current.add(pending); void pending.finally(() => pendingCheckpoints.current.delete(pending)); }
  }, []);

  useEffect(() => bridge?.onPrepareClose?.(async () => {
    if (closingRef.current || operation.current || sessionsRef.current.some(item => item.editorBusy)) return false;
    // Flush focused fields before freezing; their blur handlers may start native work.
    flushSync(() => { (window.document.activeElement as HTMLElement | null)?.blur?.(); });
    if (closingRef.current || operation.current || sessionsRef.current.some(item => item.editorBusy)) return false;
    closingRef.current = true;
    if (rootElement.current) rootElement.current.inert = true;
    flushSync(() => setClosing(true));
    while (pendingCheckpoints.current.size) await Promise.all([...pendingCheckpoints.current]);
    return !operation.current && !sessionsRef.current.some(item => item.editorBusy);
  }, () => {
    closingRef.current = false;
    if (rootElement.current) rootElement.current.inert = false;
    setClosing(false);
  }), [bridge]);

  useEffect(() => {
    const block = (event: KeyboardEvent) => { if (closingRef.current) { event.preventDefault(); event.stopImmediatePropagation(); } };
    window.addEventListener('keydown', block, true);
    return () => window.removeEventListener('keydown', block, true);
  }, []);

  const saveDocument = useCallback(async (saveAs = false): Promise<boolean> => {
    let snapshot = current.current;
    if (!snapshot || !bridge) return false;
    const commit = commitHandlers.current.get(snapshot.key);
    if (commit && !(await commit())) return false;
    snapshot = current.current;
    if (!snapshot || editorBusyRef.current || draftDirtyRef.current) return false;
    const saved = await bridge.save({ id: snapshot.id, bytes: snapshot.bytes, saveAs });
    if (!saved) return false;
    checkpointSources.current.delete(snapshot.id);
    const latest = current.current;
    if (latest?.key === snapshot.key) {
      updateDocument({ ...latest, ...saved, dirty: latest.bytes !== snapshot.bytes });
      setNotice(`Saved ${saved.name}`);
      await refreshRecent();
    }
    return true;
  }, [bridge, updateDocument, refreshRecent]);

  const runAction = useCallback(async (action: 'open' | 'save' | 'saveAs' | 'new' | 'create' | 'recover' | 'externalOpen' | 'importText', target?: string) => {
    if (closingRef.current || operation.current || editorBusyRef.current || !bridge) return;
    if (action === 'new') { setShowHome(true); return; }
    operation.current = true;
    setWorking(true);
    setError('');
    setOpenFailure(null);
    setNotice('');
    try {
      if (action === 'save' || action === 'saveAs') {
        await saveDocument(action === 'saveAs');
        return;
      }
      if (action === 'externalOpen') externalPending.current = true;
      if (sessionsRef.current.length >= 12) throw new Error('Close a document before opening another. Up to 12 documents can stay open.');
      let opened:Awaited<ReturnType<typeof bridge.open>> & {untitled?:true;recoveryDraft?:unknown}|null;
      if(action==='importText'){
        const picked=await bridge.pickDelimitedImport();if(!picked)return;
        const controller=new AbortController();importAbort.current=controller;setImportProgress({completed:0,total:0});
        const {importDelimitedWorkbook}=await import('./spreadsheetDelimited');
        const bytes=await importDelimitedWorkbook(new Uint8Array(picked.seed),new Uint8Array(picked.bytes),picked.format,{signal:controller.signal,onProgress:(completed,total)=>setImportProgress({completed,total})});
        controller.signal.throwIfAborted();
        opened=await bridge.importDocument({name:picked.name.replace(/\.(csv|tsv)$/i,'')+'.xlsx',bytes});
      }else{
        opened = action === 'externalOpen' ? await bridge.nextExternal() : action === 'recover' ? await bridge.recover(target!) : action === 'create'
        ? await bridge.create(target as 'docx' | 'xlsx' | 'pptx' | 'pdf')
        : target ? await bridge.openRecent(target) : await bridge.open();
      }
      if (action === 'externalOpen' && !opened) externalPending.current = false;
      if (opened) {
        const bytes = new Uint8Array(opened.bytes);
        // A file from disk that is not even the right kind of container never gets a tab, a ribbon
        // and a "Saved" status around an editor that can only show the engine's raw complaint.
        // Documents this app just generated (New, CSV import) are trusted as they are.
        const failure = action === 'create' || action === 'importText' ? undefined : containerFailure(opened.name, bytes);
        if (failure) {
          await bridge.close(opened.id).catch(() => { /* The host drops unopened sessions on its own. */ });
          // Stop draining the queue: another open would clear the page before it is read.
          externalPending.current = false;
          setOpenFailure({ kind: 'extract', detail: failure, name: opened.name });
          setShowHome(true);
          await refreshRecent();
          return;
        }
        const next = { ...opened, key: ++sequence.current, initialName: opened.name, initialBytes: bytes, bytes, dirty: action === 'create' || action === 'recover' || action === 'importText' };
        updateDocument(next);
        if (next.dirty) checkpoint(next);
        if (action === 'recover') setRecoveries(value => value.filter(entry => entry.id !== target));
        setNotice(action==='importText'?'Imported as literal text into a new workbook. Save to choose its location.':'');
        setShowHome(false);
        await refreshRecent();
        if (action === 'externalOpen') externalPending.current = true;
      }
    } catch (cause) {
      if(action==='importText'&&importAbort.current?.signal.aborted)setNotice('Import canceled.');
      else {
        const message = cause instanceof Error ? cause.message : String(cause);
        const kind = action === 'importText' ? undefined : classifyOpenError(message);
        if (kind) { externalPending.current = false; setOpenFailure({ kind, detail: message, name: typeof target === 'string' && target.includes('.') ? target : undefined }); setShowHome(true); }
        else setError(message);
      }
    } finally {
      if(action==='importText'){importAbort.current=undefined;setImportProgress(null);}
      bridge.setDirty(sessionsRef.current.some(item => item.dirty || item.draftDirty));
      operation.current = false;
      setWorking(false);
    }
  }, [bridge, saveDocument, updateDocument, refreshRecent, checkpoint]);

  // Undo/Redo from the Edit menu or the title bar: text fields undo through the host; editors undo their own history.
  const runHistory = useCallback((action: 'undo' | 'redo') => {
    if (closingRef.current || operation.current || editorBusyRef.current) return;
    if (draftDirtyRef.current || window.document.activeElement?.closest('dialog, .document-search, .pdf-search')) void bridge?.textHistory(action).catch(() => setError('Text history could not be updated.'));
    else if (current.current) historyHandlers.current.get(current.current.key)?.[action]();
  }, [bridge]);
  useEffect(() => bridge?.onMenuAction?.((action) => { if (closingRef.current) return; if (action === 'updates') { setUpdatesOpen(true); return; } if (action === 'undo' || action === 'redo') { runHistory(action); return; } if (action === 'externalOpen') externalPending.current = true; void runAction(action); }), [bridge, runAction, runHistory]);
  useEffect(() => { if (!working && !editorBusy && sessions.length < 12 && externalPending.current) { externalPending.current = false; void runAction('externalOpen'); } }, [working, editorBusy, sessions.length, runAction]);

  useEffect(() => {
    const handleKey = (event: KeyboardEvent) => {
      if (!(event.metaKey || event.ctrlKey) || event.altKey) return;
      const key = event.key.toLowerCase();
      if (key === 'k') { event.preventDefault(); if (!operation.current && !settingsOpen && !updatesOpen) setCommandSearch(value => !value); return; }
      if (commandSearch || settingsOpen || updatesOpen) return;
      if (key !== 'o' && key !== 's' && key !== 'n') return;
      event.preventDefault();
      void runAction(key === 'n' ? 'new' : key === 'o' ? 'open' : event.shiftKey ? 'saveAs' : 'save');
    };
    window.addEventListener('keydown', handleKey);
    return () => window.removeEventListener('keydown', handleKey);
  }, [runAction, commandSearch, settingsOpen, updatesOpen]);

  // Esc leaves Focus mode, as it does in Word. A draft being edited owns Esc first (the editor
  // cancels it), and so do dialogs, the command palette, a context menu and the search fields.
  useEffect(() => {
    if (!viewOptions.focus) return;
    const leaveFocus = (event: KeyboardEvent) => {
      if (event.key !== 'Escape' || event.repeat) return;
      if (draftDirtyRef.current || editorBusyRef.current || operation.current || closingRef.current) return;
      if (commandSearch || settingsOpen || updatesOpen || replacePrompt) return;
      if (window.document.querySelector('.context-menu, dialog[open]')) return;
      if (window.document.activeElement?.closest('.document-search, .pdf-search')) return;
      event.preventDefault();
      setViewOptions(value => ({ ...value, focus: false }));
    };
    window.addEventListener('keydown', leaveFocus);
    return () => window.removeEventListener('keydown', leaveFocus);
  }, [viewOptions.focus, commandSearch, settingsOpen, updatesOpen, replacePrompt]);

  useEffect(() => {
    window.document.title = document ? `${document.dirty || draftDirty ? '• ' : ''}${document.name} — InjOffice` : 'InjOffice';
  }, [document?.name, document?.dirty, draftDirty]);

  const changeSession = useCallback((key: number, patch: Partial<DocumentSession>) => {
    const existing = sessionsRef.current.find(item => item.key === key);
    if (!existing) return;
    const changed = { ...existing, ...patch };
    publishSessions(sessionsRef.current.map(item => item.key === key ? changed : item));
    if (patch.bytes || Object.hasOwn(patch, 'recoveryDraft')) checkpoint(changed);
  }, [publishSessions, checkpoint]);
  const closeDocument = async (key: number) => {
    if (closingRef.current || operation.current || editorBusyRef.current || !bridge) return;
    const closing = sessionsRef.current.find(item => item.key === key);
    if (!closing || closing.editorBusy) return;
    publishSessions(sessionsRef.current, key); setShowHome(false);
    operation.current = true; setWorking(true);
    try {
      if (closing.dirty || closing.draftDirty) {
        const choice = await new Promise<ReplaceChoice>(resolve => { resolveChoice.current = resolve; setReplacePrompt(true); });
        if (choice === 'cancel' || (choice === 'save' && !(await saveDocument()))) return;
      }
      await bridge.close(closing.id);
      checkpointSources.current.delete(closing.id);
      historyHandlers.current.delete(key); commitHandlers.current.delete(key); sessionViews.current.delete(key);
      setHistoryState(({ [key]: _closed, ...rest }) => rest);
      const remaining = sessionsRef.current.filter(item => item.key !== key);
      publishSessions(remaining);
      if (!remaining.length) setShowHome(true);
    } catch (cause) { setError(cause instanceof Error ? cause.message : String(cause)); }
    finally { operation.current = false; setWorking(false); }
  };

  const importFiles = async (files: File[]) => {
    if (closingRef.current || operation.current || !bridge || editorBusyRef.current) return;
    operation.current = true; setWorking(true); setError(''); setNotice('');
    const staged: DocumentSession[] = [];
    const failures: string[] = [];
    try {
      // Import the whole batch before mounting editors. Loading the first editor
      // otherwise marks the native host busy and refuses subsequent imports.
      for (const [index, file] of files.entries()) {
        if (sessionsRef.current.length + staged.length >= 12) {
          failures.push(`${files.length - index} remaining file(s): close a tab before importing more files (12-document limit).`);
          break;
        }
        try {
          if (!/\.(docx|xlsx|pptx|pdf)$/i.test(file.name) || file.size > 512 * 1024 * 1024) throw new Error('Use a DOCX, XLSX, PPTX, or PDF smaller than 512 MB.');
          const opened = await bridge.importDocument({ name: file.name, bytes: new Uint8Array(await file.arrayBuffer()) });
          if (!opened) throw new Error('The workspace was busy. Try dropping this file again.');
          const bytes = new Uint8Array(opened.bytes);
          const failure = containerFailure(opened.name, bytes);
          if (failure) { await bridge.close(opened.id).catch(() => { /* The host drops unopened sessions on its own. */ }); throw new Error(failure); }
          staged.push({ ...opened, key: ++sequence.current, initialName: opened.name, initialBytes: bytes, bytes, dirty: true });
        } catch (cause) { failures.push(`${file.name}: ${cause instanceof Error ? cause.message : String(cause)}`); }
      }
      if (staged.length) {
        publishSessions([...sessionsRef.current, ...staged], staged.at(-1)!.key);
        for (const session of staged) checkpoint(session);
        setShowHome(false);
        setNotice(`Imported ${staged.length} of ${files.length} file${files.length === 1 ? '' : 's'} as unsaved copies. Save to choose their locations.`);
      }
      if (failures.length) setError(`Could not import ${files.length - staged.length} file${files.length - staged.length === 1 ? '' : 's'}. ${failures.join(' ')}`);
    } finally { operation.current = false; setWorking(false); }
  };

  const choose = (choice: ReplaceChoice) => {
    setReplacePrompt(false);
    resolveChoice.current?.(choice);
    resolveChoice.current = null;
  };
  const busy = closing || working || editorBusy;
  const isDocx = document?.name.toLowerCase().endsWith('.docx') ?? false;
  const changeZoom = (zoom: number) => setViewOptions(value => ({ ...value, zoom: Math.max(50, Math.min(200, zoom)) }));
  const toggleFocus = () => setViewOptions(value => ({ ...value, focus: !value.focus }));
  // Office states the save state once, in the title bar. The status bar keeps messages and zoom.
  const saveStatus = document?.untitled ? 'Not saved yet' : draftDirty || document?.dirty ? 'Unsaved changes' : 'Saved';
  const canUndo = !!document && !busy && (draftDirty || (historyState[document.key]?.undo ?? true));
  const canRedo = !!document && !busy && (draftDirty || (historyState[document.key]?.redo ?? true));
  const searchLabel = `Search (${shortcutLabel('commands')})`, searchTitle = shortcutTooltip('Search commands', 'commands');

  const commands: WorkspaceCommand[] = [
    { id: 'preferences', label: 'Preferences', detail: 'Local view defaults', disabled: busy, run: () => setSettingsOpen(true) },
    { id: 'home', label: 'Go to start page', disabled: busy, run: () => setShowHome(true) },
    ...(['docx', 'xlsx', 'pptx', 'pdf'] as const).map((format, index) => ({ id: `new-${format}`, label: ['New document', 'New spreadsheet', 'New presentation', 'New blank PDF'][index], detail: format.toUpperCase(), disabled: busy || !bridge, run: () => { void runAction('create', format); } })),
    {id:'importText',label:'Import CSV or TSV',detail:'New workbook · literal text',disabled:busy||!bridge,run:()=>{void runAction('importText')}},
    { id: 'open', label: 'Open file', detail: 'Ctrl / ⌘ O', disabled: busy || !bridge, run: () => { void runAction('open'); } },
    { id: 'save', label: 'Save document', detail: 'Ctrl / ⌘ S', disabled: busy || !document, run: () => { void runAction('save'); } },
    { id: 'save-as', label: 'Save document as…', disabled: busy || !document, run: () => { void runAction('saveAs'); } },
    { id: 'close', label: 'Close document', detail: document?.name, disabled: busy || !document, run: () => { if (document) void closeDocument(document.key); } },
    { id: 'focus', label: viewOptions.focus ? 'Exit focus mode' : 'Enter focus mode', disabled: !document, run: () => { setShowHome(false); toggleFocus(); } },
    { id: 'updates', label: 'App updates', run: () => setUpdatesOpen(true) },
    { id: 'zoom-reset', label: 'Reset document zoom to 100%', disabled: !document, run: () => changeZoom(100) },
    { id: 'zoom-in', label: 'Zoom in', disabled: !document || viewOptions.zoom >= 200, run: () => changeZoom(viewOptions.zoom + 10) },
    { id: 'zoom-out', label: 'Zoom out', disabled: !document || viewOptions.zoom <= 50, run: () => changeZoom(viewOptions.zoom - 10) },
    ...(isDocx ? [{ id: 'outline', label: viewOptions.navigation ? 'Hide document outline' : 'Show document outline', run: () => setViewOptions(value => ({ ...value, navigation: !value.navigation })) }] : []),
    ...sessions.map(item => ({ id: `tab-${item.key}`, label: `Switch to ${item.name}`, detail: 'Open document', disabled: busy, run: () => { publishSessions(sessionsRef.current, item.key); setShowHome(false); } })),
  ];

  // Office keeps New/Open/Save/Close in the File backstage; every editor's ribbon renders these
  // groups in its File tab (before its own Export group) through WorkspaceFileGroupsContext.
  const fileGroups: WorkspaceFileGroups = {
    before: [
      { id: 'workspace-home', label: 'Start', children: <RibbonButton icon="home" label="Home" title="Start page: create, open recent or recover" disabled={busy} onClick={() => setShowHome(true)} /> },
      { id: 'workspace-new', label: 'New', children: <RibbonRows>
        <div><RibbonButton icon="newDocument" label="Document" title="New document (DOCX)" disabled={busy || !bridge} onClick={() => void runAction('create', 'docx')} /><RibbonButton icon="newDocument" label="Spreadsheet" title="New spreadsheet (XLSX)" disabled={busy || !bridge} onClick={() => void runAction('create', 'xlsx')} /></div>
        <div><RibbonButton icon="newDocument" label="Presentation" title="New presentation (PPTX)" disabled={busy || !bridge} onClick={() => void runAction('create', 'pptx')} /><RibbonButton icon="newDocument" label="PDF" title="New blank PDF" disabled={busy || !bridge} onClick={() => void runAction('create', 'pdf')} /></div>
      </RibbonRows> },
      { id: 'workspace-open-save', label: 'Open & Save', children: <>
        <RibbonButton icon="open" label="Open" shortcut="open" disabled={busy || !bridge} onClick={() => void runAction('open')} />
        <RibbonButton icon="save" label="Save" shortcut="save" disabled={busy || !document || (!document.dirty && !draftDirty)} onClick={() => void runAction('save')} />
        <RibbonButton icon="saveAs" label="Save as…" shortcut="saveAs" disabled={busy || !document} onClick={() => void runAction('saveAs')} />
      </> },
    ],
    after: [
      { id: 'workspace-close', label: 'Close', children: <RibbonButton icon="closeDocument" label="Close document" disabled={busy || !document} onClick={() => { if (document) void closeDocument(document.key); }} /> },
      { id: 'workspace-app', label: 'InjOffice', children: <>
        <RibbonButton icon="appUpdate" label="Updates" title="App updates" onClick={() => setUpdatesOpen(true)} />
        <RibbonButton icon="settings" label="Preferences" title="Preferences: local view defaults and appearance" disabled={busy} onClick={() => setSettingsOpen(true)} />
      </> },
    ],
  };

  return (
    <div ref={rootElement} inert={closing ? true : undefined} className={`desktop-app${macChrome ? ' platform-mac' : ''}${viewOptions.focus ? ' is-focused' : ''}`} onDragOver={event => { if (event.dataTransfer.types.includes('Files')) event.preventDefault(); }} onDrop={event => { if (event.defaultPrevented || !event.dataTransfer.files.length) return; event.preventDefault(); void importFiles(Array.from(event.dataTransfer.files)); }}>
      {!showHome && document && <header className="app-titlebar">
        {/* Office's title bar: Quick Access (Save, Undo, Redo) · document name and state · Search · view tools. File, New, Open and Save as… live in the ribbon's File tab. */}
        <div className="titlebar-quick-access" role="toolbar" aria-label="Quick access">
          <RibbonButton className="titlebar-button" icon="save" label="Save" shortcut="save" labelHidden disabled={busy || (!document.dirty && !draftDirty)} onClick={() => void runAction('save')} />
          <RibbonButton className="titlebar-button" icon="undo" label="Undo" shortcut="undo" labelHidden disabled={!canUndo} onClick={() => runHistory('undo')} />
          <RibbonButton className="titlebar-button" icon="redo" label="Redo" shortcut="redo" labelHidden disabled={!canRedo} onClick={() => runHistory('redo')} />
        </div>
        <div className="titlebar-centre">
          <div className="title-document">
            <span className={`document-format format-${document.name.split('.').pop()?.toLowerCase()}`}>{document.name.split('.').pop()?.toUpperCase()}</span>
            <span className="document-name" title={document.name}>{document.name}</span>
            <span className="title-save-status" role="status">{(document.dirty || draftDirty) && <span className="dirty-indicator" aria-hidden="true" />}{saveStatus}</span>
          </div>
          <button className="titlebar-search" disabled={busy} onClick={() => setCommandSearch(true)} title={searchTitle}><RibbonIcon name="find" /><span>{searchLabel}</span></button>
        </div>
        <div className="titlebar-tools" role="toolbar" aria-label="View">
          {isDocx && <RibbonButton className="titlebar-button" icon="sidebar" label={viewOptions.navigation ? 'Hide document outline' : 'Show document outline'} labelHidden aria-pressed={viewOptions.navigation} onClick={() => setViewOptions(value => ({ ...value, navigation: !value.navigation }))} />}
          <RibbonButton className="titlebar-button" icon="focus" label={viewOptions.focus ? 'Exit focus mode' : 'Focus mode'} labelHidden aria-pressed={viewOptions.focus} onClick={toggleFocus} />
          <RibbonButton className="titlebar-button" icon="appUpdate" label="App updates" labelHidden onClick={() => setUpdatesOpen(true)} />
        </div>
      </header>}

      {importProgress&&<div className="app-import-progress" role="status"><span>{importProgress.total?`Importing spreadsheet: ${importProgress.completed.toLocaleString()} of ${importProgress.total.toLocaleString()} fields`:'Preparing spreadsheet import…'}</span><button onClick={()=>importAbort.current?.abort()}>Cancel import</button></div>}
      {error && <div className="app-error" role="alert"><span>{error}</span><button onClick={() => setError('')} aria-label="Dismiss error">×</button></div>}

      {showHome && (openFailure
        ? <OpenError name={openFailure.name} kind={openFailure.kind} detail={openFailure.detail} busy={busy} onOpen={() => { setOpenFailure(null); void runAction('open'); }} onHome={() => setOpenFailure(null)} />
        : <StartPage recoveries={recoveries} onRecover={id => void runAction('recover', id)} onDiscardRecovery={id => { void bridge?.discardRecovery(id).then(() => setRecoveries(value => value.filter(item => item.id !== id))).catch(() => setError('Recovery copy could not be discarded.')); }} openDocuments={sessions.map(item => ({ key: item.key, name: item.name }))} onSelectDocument={key => { publishSessions(sessionsRef.current, key); setShowHome(false); }} onSearch={() => setCommandSearch(true)} searchLabel={searchLabel} searchTitle={searchTitle} onUpdates={() => setUpdatesOpen(true)} recentFiles={recentFiles} busy={busy} available={Boolean(bridge)} onImportText={()=>void runAction('importText')} onOpen={() => void runAction('open')} onCreate={format => void runAction('create', format)} onOpenRecent={id => void runAction('open', id)} onRemoveRecent={id => void removeRecent(id)} currentName={document?.name} onResume={document ? () => setShowHome(false) : undefined} />)}

      {sessions.length > 0 && <nav className="document-tabs" aria-label="Open documents" hidden={showHome}>
        {sessions.map(item => <div className={`document-tab${item.key === document?.key ? ' active' : ''}`} key={item.key}>
          <button aria-current={item.key === document?.key ? 'page' : undefined} disabled={busy} onClick={() => publishSessions(sessionsRef.current, item.key)} title={item.name}><span className={`tab-format format-${item.name.split('.').pop()}`}>{item.name.split('.').pop()?.toUpperCase()}</span>{item.dirty || item.draftDirty ? '• ' : ''}{item.name}</button>
          <button disabled={busy} aria-label={`Close ${item.name}`} title={`Close ${item.name}`} onClick={() => void closeDocument(item.key)}><RibbonIcon name="close" /></button>
        </div>)}
      </nav>}
      <WorkspaceFileGroupsContext value={fileGroups}>{sessions.map(item => <main key={item.key} className="editor-workspace" hidden={showHome || item.key !== document?.key} aria-label={item.name} aria-busy={item.editorBusy}>
        <div className="editor-content" inert={working || item.key !== document?.key ? true : undefined}>
          <SessionEditor session={item} registerSessionHistory={registerSessionHistory} registerSessionCommit={registerSessionCommit} onSessionChange={changeSession} viewOptions={sessionViews.current.get(item.key) ?? viewOptions} />
        </div>
      </main>)}</WorkspaceFileGroupsContext>

      <footer className="app-status" hidden={showHome}><span className="status-message" role="status">{busy ? 'Working…' : draftDirty ? 'Draft changes · Save applies your edits.' : notice}</span><div className="status-view-controls"><div className="status-zoom"><button disabled={!document || viewOptions.zoom <= 50} aria-label="Zoom out" onClick={() => changeZoom(viewOptions.zoom - 10)}>−</button><input type="range" aria-label="Document zoom" min="50" max="200" step="5" disabled={!document} value={viewOptions.zoom} onChange={event => changeZoom(Number(event.target.value))} /><button disabled={!document || viewOptions.zoom >= 200} aria-label="Zoom in" onClick={() => changeZoom(viewOptions.zoom + 10)}>+</button><output>{viewOptions.zoom}%</output></div></div></footer>

      <UpdateNotice onOpen={() => setUpdatesOpen(true)} />
      {updatesOpen && <UpdatesDialog onClose={() => setUpdatesOpen(false)} />}
      {settingsOpen && <PreferencesDialog value={preferences} onClose={() => setSettingsOpen(false)} onSave={value => { try { writePreferences(value); setPreferences(value); setSettingsOpen(false); setNotice('Preferences saved on this device.'); } catch { setError('Preferences could not be saved on this device.'); } }} />}
      {commandSearch && <CommandPalette commands={commands} onClose={() => setCommandSearch(false)} />}
      {replacePrompt && <div className="modal-backdrop"><section className="save-dialog" role="alertdialog" aria-modal="true" aria-labelledby="save-dialog-title" aria-describedby="save-dialog-description" onKeyDown={(event) => {
        if (event.key === 'Escape') choose('cancel');
        if (event.key !== 'Tab') return;
        const buttons = event.currentTarget.querySelectorAll<HTMLButtonElement>('button');
        const first = buttons[0];
        const last = buttons[buttons.length - 1];
        if (event.shiftKey && window.document.activeElement === first) { event.preventDefault(); last?.focus(); }
        if (!event.shiftKey && window.document.activeElement === last) { event.preventDefault(); first?.focus(); }
      }}>
        <h2 id="save-dialog-title">Save your changes?</h2>
        <p id="save-dialog-description">Save changes to <strong>{document?.name}</strong> before closing this tab.</p>
        <div className="dialog-actions"><button onClick={() => choose('discard')}>Discard changes</button><button autoFocus onClick={() => choose('cancel')}>Cancel</button><button className="primary-button" onClick={() => choose('save')}>Save changes</button></div>
      </section></div>}
    </div>
  );
}
