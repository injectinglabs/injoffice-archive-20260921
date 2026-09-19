export {};

declare global {
  interface DesktopUpdateState {
    status: 'disabled' | 'idle' | 'checking' | 'available' | 'downloading' | 'downloaded' | 'installing' | 'not-available' | 'error';
    appVersion: string;
    autoCheck: boolean;
    version?: string;
    releaseNotes?: string;
    percent?: number;
    message?: string;
  }
  interface RecentFile { id: string; name: string; path: string; updatedAt: number }
  interface Window {
    injDesktop?: {
      getUpdateState(): Promise<DesktopUpdateState>;
      checkForUpdates(): Promise<DesktopUpdateState>;
      downloadUpdate(): Promise<DesktopUpdateState>;
      installUpdate(): Promise<DesktopUpdateState>;
      setAutomaticUpdates(enabled: boolean): Promise<DesktopUpdateState>;
      onUpdateState(callback: (state: DesktopUpdateState) => void): () => void;
      textHistory(direction: 'undo' | 'redo'): Promise<void>;
      nextExternal(): Promise<{ id: string; name: string; bytes: Uint8Array } | null>;
      open(): Promise<{ id: string; name: string; bytes: Uint8Array } | null>;
      create(format: 'docx' | 'xlsx' | 'pptx' | 'pdf'): Promise<{ id: string; name: string; bytes: Uint8Array; untitled: true } | null>;
      pickDelimitedImport():Promise<{name:string;format:'csv'|'tsv';bytes:Uint8Array;seed:Uint8Array}|null>;
      pickAsset(kind: 'pdf' | 'image'): Promise<{ name: string; bytes: Uint8Array } | null>;
      exportDocxPdf(input:{requestId:string;bytes:Uint8Array;revision:string;mode?:'original'|'preview'}):Promise<Uint8Array>;
      cancelDocxPdf(requestId:string):Promise<boolean>;
      onPdfExportProgress?(callback:(value:{requestId:string;stage:'reading'|'layout'|'outlining'|'painting'})=>void):()=>void;
      replacePdfText(input: { bytes: Uint8Array; revision: string; page: number; rect: [number, number, number, number]; oldText: string; newText: string }): Promise<Uint8Array>;
      exportBytes(input: { name: string; bytes: Uint8Array }): Promise<{ name: string } | null>;
      importDocument(input: { name: string; bytes: Uint8Array }): Promise<{ id: string; name: string; bytes: Uint8Array; untitled: true } | null>;
      checkpoint(id: string, bytes: Uint8Array | null, draft?: unknown, revision?: string): Promise<void>;
      recovery(): Promise<Array<{ id: string; name: string; updatedAt: number; size: number }>>;
      recover(id: string): Promise<{ id: string; name: string; bytes: Uint8Array; untitled: true; recoveryDraft?: unknown } | null>;
      discardRecovery(id: string): Promise<void>;
      close(id: string): Promise<void>;
      recent(): Promise<RecentFile[]>;
      openRecent(id: string): Promise<{ id: string; name: string; bytes: Uint8Array } | null>;
      removeRecent(id: string): Promise<void>;
      save(input: { id: string; bytes: Uint8Array; saveAs?: boolean }): Promise<{ id: string; name: string; untitled: false } | null>;
      setDirty(dirty: boolean): void;
      setBusy?(busy: boolean): void;
      onPrepareClose?(prepare: () => Promise<boolean>, cancel: () => void): () => void;
      onMenuAction?(callback: (action: 'new' | 'open' | 'save' | 'saveAs' | 'importText' | 'externalOpen' | 'updates' | 'undo' | 'redo') => void): () => void;
    };
  }
}
