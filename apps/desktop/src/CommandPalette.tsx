import { useEffect, useRef, useState } from 'react';
export type WorkspaceCommand = { id: string; label: string; detail?: string; disabled?: boolean; run(): void };
export default function CommandPalette({ commands, onClose }: { commands: WorkspaceCommand[]; onClose(): void }) {
  const dialog = useRef<HTMLDialogElement>(null);
  const [query, setQuery] = useState('');
  const [selected, setSelected] = useState(0);
  const results = commands.filter(command => `${command.label} ${command.detail ?? ''}`.toLowerCase().includes(query.toLowerCase().trim()));
  const active = Math.min(selected, Math.max(0, results.length - 1));
  useEffect(() => { dialog.current?.showModal(); }, []);
  useEffect(() => { dialog.current?.querySelector('[aria-selected="true"]')?.scrollIntoView({ block: 'nearest' }); }, [active, query]);
  const execute = (command: WorkspaceCommand) => { if (!command.disabled) { onClose(); command.run(); } };
  return <dialog ref={dialog} className="command-palette" aria-labelledby="command-palette-title" onCancel={event => { event.preventDefault(); onClose(); }} onClick={event => { if (event.target === dialog.current) { const rect = dialog.current.getBoundingClientRect(); if (event.clientX < rect.left || event.clientX > rect.right || event.clientY < rect.top || event.clientY > rect.bottom) onClose(); } }}>
    <header><h2 id="command-palette-title">Search commands</h2><button onClick={onClose} aria-label="Close command search">Esc</button></header>
    <input autoFocus type="search" placeholder="New document, save, zoom…" value={query} aria-label="Search workspace commands" role="combobox" aria-expanded="true" aria-controls="command-results" aria-activedescendant={results[active] ? `command-${results[active].id}` : undefined} onChange={event => { setQuery(event.target.value); setSelected(0); }} onKeyDown={event => {
      if (event.key === 'ArrowDown' || event.key === 'ArrowUp') { event.preventDefault(); setSelected((active + (event.key === 'ArrowDown' ? 1 : -1) + results.length) % (results.length || 1)); }
      if (event.key === 'Enter' && results[active]) { event.preventDefault(); execute(results[active]); }
    }} />
    <div id="command-results" role="listbox" aria-label="Workspace commands">{results.map((command, index) => <div key={command.id} id={`command-${command.id}`} role="option" aria-selected={index === active} aria-disabled={command.disabled || undefined} onMouseEnter={() => setSelected(index)} onMouseDown={event => event.preventDefault()} onClick={() => execute(command)}><span>{command.label}</span>{command.detail && <small>{command.detail}</small>}</div>)}</div>
    {!results.length && <p className="command-empty">No matching workspace commands.</p>}
  </dialog>;
}
