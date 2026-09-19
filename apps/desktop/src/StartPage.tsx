import { useMemo, useState } from 'react';
import logo from '../../../logo.png';
import './start-page.css';

export interface StartPageProps {
  recentFiles: Array<{ id: string; name: string; path: string; updatedAt: number }>;
  busy: boolean;
  available: boolean;
  onOpen: () => void;
  onUpdates?: () => void;
  onImportText?:()=>void;
  onCreate: (format: 'docx' | 'xlsx' | 'pptx' | 'pdf') => void;
  onOpenRecent: (id: string) => void;
  onRemoveRecent: (id: string) => void;
  onResume?: () => void;
  currentName?: string;
}

const formats = [
  { extension: 'docx', name: 'Documents', letter: 'W', detail: 'DOCX', color: '#2459ad' },
  { extension: 'xlsx', name: 'Spreadsheets', letter: 'S', detail: 'XLSX', color: '#28764f' },
  { extension: 'pptx', name: 'Presentations', letter: 'P', detail: 'PPTX', color: '#b95b31' },
  { extension: 'pdf', name: 'PDF files', letter: 'P', detail: 'PDF', color: '#b8494f' },
] as const;

function Icon({ name }: { name: 'home' | 'open' | 'search' | 'remove' | 'return' | 'clock' }) {
  const paths = {
    home: 'm3 10 9-7 9 7M5 9v12h5v-7h4v7h5V9',
    open: 'M3 7V5h6l2 2h10v3M3 10h19l-3 10H3V10Z',
    search: 'm16 16 5 5M18 10a8 8 0 1 1-16 0 8 8 0 0 1 16 0Z',
    remove: 'm6 6 12 12M18 6 6 18',
    return: 'm8 5-5 5 5 5M3 10h12a5 5 0 0 1 0 10h-3',
    clock: 'M12 7v5l3 2M22 12a10 10 0 1 1-20 0 10 10 0 0 1 20 0Z',
  };
  return <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d={paths[name]} /></svg>;
}

function formatFor(name: string) {
  return formats.find(format => name.toLowerCase().endsWith(`.${format.extension}`));
}

function FileBadge({ name, large = false }: { name: string; large?: boolean }) {
  const format = formatFor(name);
  return <span className={`start-file-badge${large ? ' start-file-badge-large' : ''}`} style={{ color: format?.color ?? '#687386' }} aria-hidden="true"><svg viewBox="0 0 32 38" fill="none"><path d="M5 1h15l8 8v27H5V1Z" fill="currentColor" opacity=".09" /><path d="M5 1h15l8 8v27H5V1Z" stroke="currentColor" strokeWidth="1.3" /><path d="M20 1v8h8" stroke="currentColor" strokeWidth="1.3" /></svg><span>{format?.detail ?? 'FILE'}</span></span>;
}

function recentDate(timestamp: number) {
  const date = new Date(timestamp);
  if (!Number.isFinite(date.getTime())) return 'Unknown date';
  const now = new Date();
  const sameDay = date.toDateString() === now.toDateString();
  return sameDay ? `Today, ${date.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' })}` : date.toLocaleDateString([], { month: 'short', day: 'numeric', ...(date.getFullYear() !== now.getFullYear() ? { year: 'numeric' } : {}) });
}

export default function StartPage({ recentFiles, busy, available, onOpen, onUpdates, onImportText, onCreate, onOpenRecent, onRemoveRecent, onResume, currentName }: StartPageProps) {
  const [query, setQuery] = useState('');
  const [filter, setFilter] = useState('all');
  const visibleFiles = useMemo(() => recentFiles.filter(file => {
    const matchesType = filter === 'all' || file.name.toLowerCase().endsWith(`.${filter}`);
    const search = query.trim().toLowerCase();
    return matchesType && (!search || `${file.name} ${file.path}`.toLowerCase().includes(search));
  }).sort((a, b) => b.updatedAt - a.updatedAt), [recentFiles, query, filter]);
  const disabled = busy || !available;

  return <div className="start-page" aria-busy={busy}>
    <aside className="start-sidebar" aria-label="Start navigation">
      <div className="start-brand"><img src={logo} alt="" /><span>InjOffice</span></div>
      <nav className="start-nav">
        <span className="start-nav-home" aria-current="page"><Icon name="home" />Home</span>
        <button onClick={onOpen} disabled={disabled}><Icon name="open" />Open file</button>
        {onImportText&&<button onClick={onImportText} disabled={disabled} title="Create a new spreadsheet with CSV/TSV fields imported as literal text"><Icon name="open" />Import CSV / TSV</button>}
        {onResume && <button onClick={onResume} disabled={busy} title={currentName ? `Return to ${currentName}` : 'Return to document'}><Icon name="return" /><span>Return to document</span></button>}
        {onUpdates && <button onClick={onUpdates}>App updates</button>}
      </nav>
      {currentName && onResume && <div className="start-current"><span>Currently open</span><button onClick={onResume} disabled={busy} title={currentName}><FileBadge name={currentName} /><span>{currentName}</span></button></div>}
    </aside>

    <main className="start-main">
      <div className="start-main-inner">
        <section className="start-intro" aria-labelledby="start-title">
          <div><h1 id="start-title">Your next idea starts here.</h1><p>Create something new, or pick up where you left off.</p></div>
          <div className="start-open-existing"><button className="start-open-primary" onClick={onOpen} disabled={disabled}><Icon name="open" />Open existing file…</button><span>DOCX, XLSX, PPTX &amp; PDF</span></div>
        </section>
        <section className="start-create" aria-labelledby="start-create-title">
          <h2 id="start-create-title">Create new</h2>
          <div className="start-create-grid">
            {([
              { format: 'docx', label: 'Blank document' },
              { format: 'xlsx', label: 'Blank spreadsheet' },
              { format: 'pptx', label: 'Blank presentation' },
              { format: 'pdf', label: 'Blank PDF' },
            ] as const).map(item => <button key={item.format} className={`start-create-card start-create-${item.format}`} onClick={() => onCreate(item.format)} disabled={disabled} aria-label={`Create ${item.label.toLowerCase()}`}>
              <span className="start-create-preview" aria-hidden="true"><span className="start-create-paper">
                {item.format === 'docx' && <><span className="start-document-line" /><span className="start-document-line" /><span className="start-document-line" /><span className="start-document-line" /></>}
                {item.format === 'xlsx' && <span className="start-sheet-grid"><span /></span>}
                {item.format === 'pptx' && <><span className="start-slide-title" /><span className="start-slide-subtitle" /></>}
              </span><span className="start-create-plus">+</span></span>
              <span className="start-create-caption"><strong>{item.label}</strong><small>{item.format.toUpperCase()}</small></span>
            </button>)}
          </div>
        </section>
        {!available && <p className="start-unavailable" role="status">Launch InjOffice on your computer to create and open local files.</p>}

        <section className="start-recents" aria-labelledby="start-recents-title">
          <div className="start-recents-heading"><h2 id="start-recents-title">Recent files</h2><div className="start-recent-tools"><label className="start-search"><Icon name="search" /><input type="search" aria-label="Search recent files" placeholder="Search files" value={query} onChange={event => setQuery(event.target.value)} /></label><select aria-label="Filter recent files by type" value={filter} onChange={event => setFilter(event.target.value)}><option value="all">All files</option>{formats.map(format => <option key={format.extension} value={format.extension}>{format.name}</option>)}</select></div></div>
          {visibleFiles.length > 0 ? <div className="start-recent-list">
            <div className="start-list-labels" aria-hidden="true"><span>Name</span><span>Last used</span><span /></div>
            <ul>{visibleFiles.map(file => <li key={file.id} className="start-recent-row"><button className="start-recent-open" disabled={disabled} onClick={() => onOpenRecent(file.id)} aria-label={`Open ${file.name}`} title={file.path}><FileBadge name={file.name} /><span className="start-file-description"><strong>{file.name}</strong><small>{file.path}</small></span><time dateTime={Number.isFinite(new Date(file.updatedAt).getTime()) ? new Date(file.updatedAt).toISOString() : undefined}>{recentDate(file.updatedAt)}</time></button><button className="start-recent-remove" disabled={busy} onClick={() => onRemoveRecent(file.id)} aria-label={`Remove ${file.name} from recent files`} title="Remove from recent files"><Icon name="remove" /></button></li>)}</ul>
          </div> : <div className="start-empty" role="status"><Icon name={recentFiles.length ? 'search' : 'clock'} /><h3>{recentFiles.length ? 'No matching files' : 'Your recent files will appear here'}</h3><p>{recentFiles.length ? 'Try a different name or file type.' : 'Files you open or save will appear here, ready for next time.'}</p>{recentFiles.length > 0 && <button onClick={() => { setQuery(''); setFilter('all'); }}>Clear filters</button>}</div>}
        </section>
      </div>
    </main>
  </div>;
}
