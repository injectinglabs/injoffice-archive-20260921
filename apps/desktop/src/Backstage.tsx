import { useEffect, useRef, type ReactNode } from 'react';

export const backstagePages = ['Home', 'New', 'Open', 'Info', 'Save', 'Save as', 'Export', 'Close', 'Preferences', 'Updates'] as const;
export type BackstagePage = typeof backstagePages[number];
export interface BackstageProps {
  page: BackstagePage;
  onPage(page: BackstagePage): void;
  onBack(): void;
  name: string;
  path?: string;
  size: number;
  dirty: boolean;
  busy: boolean;
  available: boolean;
  recentFiles: Array<{ id: string; name: string; path: string }>;
  onCreate(format: 'docx' | 'xlsx' | 'pptx' | 'pdf'): void;
  onOpen(id?: string): void;
  onSave(saveAs: boolean): void;
  onCloseDocument(): void;
  onPreferences(): void;
  onUpdates(): void;
  exportContent?: ReactNode;
}

const capabilities: Record<string, string> = {
  docx: 'Document text and supported formatting, tables, images and page layout. PDF export uses outlined text; unsupported content may not render faithfully.',
  xlsx: 'Cell values, formulas, supported formatting and sheet operations. Delimited export saves one sheet and does not retain workbook formatting.',
  pptx: 'Supported slide text, shapes and arrangement. SVG export includes supported objects on the current slide; retain the PPTX for other content.',
  pdf: 'Text, annotations, forms and page operations. Existing text replacement requires a supported original font. Form PDFs cannot be split or imported.',
};

/** Full-window File surface. Editors stay mounted so opening File never loses a draft. */
export default function Backstage(props: BackstageProps) {
  const { page, onPage, onBack, busy, available } = props;
  const root = useRef<HTMLElement>(null);
  const back = useRef<HTMLButtonElement>(null);
  useEffect(() => {
    const previous = document.activeElement as HTMLElement | null;
    const app = document.querySelector<HTMLElement>('.desktop-app');
    const wasInert = app?.inert;
    if (app) app.inert = true;
    back.current?.focus();
    return () => { if (app) app.inert = wasInert ?? false; if (previous?.isConnected) previous.focus(); };
  }, []);
  const disabled = busy || !available;
  const format = props.name.split('.').pop()?.toLowerCase() ?? '';
  const cards = <div className="backstage-cards">{(['docx', 'xlsx', 'pptx', 'pdf'] as const).map((kind, index) => <button key={kind} disabled={disabled} onClick={() => props.onCreate(kind)}>
    <span className={`backstage-paper backstage-paper-${kind}`} aria-hidden="true"><span />{kind === 'docx' && <><span /><span /><span /></>}</span>
    <strong>{['Blank document', 'Blank spreadsheet', 'Blank presentation', 'Blank PDF'][index]}</strong><small>{kind.toUpperCase()}</small>
  </button>)}</div>;
  const recents = <section><h2>Recent files</h2>{props.recentFiles.length ? <ul className="backstage-recents">{props.recentFiles.map(file => <li key={file.id}><button disabled={disabled} onClick={() => props.onOpen(file.id)}><strong>{file.name}</strong><small>{file.path}</small></button></li>)}</ul> : <p>Files you open or save will appear here.</p>}</section>;
  return <section ref={root} className="backstage" role="dialog" aria-modal="true" aria-label="File backstage" onKeyDown={event => {
    if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); onBack(); }
    if (event.key === 'Tab') {
      const controls = Array.from(root.current?.querySelectorAll<HTMLElement>('button:not(:disabled), select:not(:disabled), input:not(:disabled), summary, [tabindex="0"]') ?? []).filter(node => node.getClientRects().length);
      const first = controls[0], last = controls.at(-1);
      if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last?.focus(); }
      if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first?.focus(); }
    }
  }}>
    <aside className="backstage-rail"><button ref={back} className="backstage-back" onClick={onBack} aria-label="Back to document" title="Back to document (Esc)">←</button><nav aria-label="File">{backstagePages.map(entry => <button key={entry} aria-current={page === entry ? 'page' : undefined} onClick={() => onPage(entry)}>{entry}</button>)}</nav></aside>
    <div className="backstage-content"><header><h1>{page}</h1><p>{props.name}{props.dirty ? ' · Unsaved changes' : ''}</p></header>
      {page === 'Home' && <><h2>Create new</h2>{cards}{recents}</>}
      {page === 'New' && cards}
      {page === 'Open' && <><button disabled={disabled} onClick={() => props.onOpen()}>Browse files…</button>{recents}</>}
      {page === 'Info' && <><h2>Document properties</h2><dl className="backstage-properties"><dt>Location</dt><dd>{props.path ?? 'Not saved to a known location'}</dd><dt>Size (current document)</dt><dd>{props.size.toLocaleString()} bytes</dd><dt>Modified on disk</dt><dd>Unavailable — the desktop host does not expose file modification times.</dd></dl><h2>Engine capabilities</h2><p>{capabilities[format] ?? 'Capabilities are not available for this file type.'}</p><p>Individual commands are enabled only when supported by the current document and editor state.</p></>}
      {(page === 'Save' || page === 'Save as') && <><h2>{page === 'Save' ? 'Save your changes' : 'Save a copy'}</h2><p>{page === 'Save' ? 'Save the current document to your computer.' : 'Choose a name and location for this document.'}</p><button disabled={disabled} onClick={() => props.onSave(page === 'Save as')}>{page === 'Save' ? 'Save document' : 'Choose location…'}</button></>}
      {page === 'Export' && <><h2>Export this document</h2><div className="backstage-exports" onClick={event => { if ((event.target as HTMLElement).closest('button:not(:disabled)')) onBack(); }}>{props.exportContent || <p>No export is available for this document.</p>}</div><p>Printing is not available in this desktop editor. Export a supported format to print from another application.</p></>}
      {page === 'Close' && <><h2>Close this document</h2><p>You will be asked to save any unsaved changes.</p><button disabled={busy} onClick={props.onCloseDocument}>Close document</button></>}
      {page === 'Preferences' && <><h2>Make yourself at home</h2><p>Choose the appearance and default document view for this device.</p><button disabled={busy} onClick={props.onPreferences}>Open preferences</button><h2>Account</h2><p>InjOffice works with local files. Account sign-in is not available.</p><button disabled title="Account sign-in is not supported">Sign in</button></>}
      {page === 'Updates' && <><h2>Keep InjOffice up to date</h2><p>See the installed version, update availability and automatic update settings.</p><button onClick={props.onUpdates}>Manage updates</button></>}
    </div>
  </section>;
}
