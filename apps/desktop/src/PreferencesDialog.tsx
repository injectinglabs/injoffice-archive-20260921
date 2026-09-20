import { useEffect, useRef, useState } from 'react';
import { defaultPreferences, type Preferences, type ThemePreference } from './preferences';
export default function PreferencesDialog({ value, onSave, onClose }: { value: Preferences; onSave(value: Preferences): void; onClose(): void }) {
  const dialog = useRef<HTMLDialogElement>(null);
  const [draft, setDraft] = useState(value);
  useEffect(() => { dialog.current?.showModal(); }, []);
  return <dialog ref={dialog} className="preferences-dialog" aria-labelledby="preferences-title" onCancel={event => { event.preventDefault(); onClose(); }}>
    <h2 id="preferences-title">Preferences</h2>
    <p>Saved on this device. View defaults apply to newly opened documents; appearance applies at once.</p>
    <label>Default document zoom<select autoFocus value={draft.defaultZoom} onChange={event => setDraft({ ...draft, defaultZoom: Number(event.target.value) })}>{[50,75,90,100,110,125,150,175,200].map(zoom => <option key={zoom} value={zoom}>{zoom}%</option>)}</select></label>
    <label>Appearance<select value={draft.theme} onChange={event => setDraft({ ...draft, theme: event.target.value as ThemePreference })}><option value="system">Match system</option><option value="light">Light</option><option value="dark">Dark</option></select></label>
    <label className="preferences-checkbox"><input type="checkbox" checked={draft.showNavigation} onChange={event => setDraft({ ...draft, showNavigation: event.target.checked })} />Show the document navigation panel</label>
    <p className="preferences-shortcuts">Keyboard shortcuts: Ctrl / ⌘ O to open, S to save, Shift S to save as, N for the start page, and K to search commands.</p>
    <div className="dialog-actions"><button onClick={() => setDraft({ ...defaultPreferences })}>Restore defaults</button><button onClick={onClose}>Cancel</button><button className="primary-button" onClick={() => onSave(draft)}>Save preferences</button></div>
  </dialog>;
}
