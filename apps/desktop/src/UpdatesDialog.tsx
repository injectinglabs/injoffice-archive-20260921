import { useEffect, useRef, useState } from 'react';
import logo from '../../../logo.png';
import './updates-dialog.css';

const labels: Record<DesktopUpdateState['status'], string> = {
  disabled: 'Updates unavailable', idle: 'Keep InjOffice up to date', checking: 'Checking for updates…',
  available: 'A new version is available', downloading: 'Downloading your update…',
  downloaded: 'Ready to restart', installing: 'Preparing to restart…', 'not-available': 'You’re up to date', error: 'Update interrupted',
};

export function UpdateNotice({ onOpen }: { onOpen(): void }) {
  const [state, setState] = useState<DesktopUpdateState | null>(null);
  const [dismissed, setDismissed] = useState('');
  const bridge = window.injDesktop;
  useEffect(() => {
    if (!bridge) return;
    let active = true, received = false;
    const unsubscribe = bridge.onUpdateState(value => { received = true; if (active) setState(value); });
    void bridge.getUpdateState().then(value => { if (active && !received) setState(value); }).catch(() => {});
    return () => { active = false; unsubscribe(); };
  }, [bridge]);
  const key = `${state?.version}:${state?.status}`;
  if (!state || !['available', 'downloaded'].includes(state.status) || dismissed === key) return null;
  return <div className="update-notice" role="status"><span>{state.status === 'downloaded' ? 'Your InjOffice update is ready.' : `InjOffice ${state.version} is available.`}</span><button onClick={onOpen}>View update</button><button aria-label="Dismiss update notice" onClick={() => setDismissed(key)}>×</button></div>;
}

export default function UpdatesDialog({ onClose }: { onClose(): void }) {
  const dialog = useRef<HTMLDialogElement>(null);
  const [state, setState] = useState<DesktopUpdateState | null>(null);
  const [pending, setPending] = useState(false);
  const [error, setError] = useState('');
  const mounted = useRef(false);
  const generation = useRef(0);
  const eventVersion = useRef(0);
  const bridge = window.injDesktop;
  useEffect(() => {
    dialog.current?.showModal();
    mounted.current = true;
    const cycle = ++generation.current;
    const unsubscribe = bridge?.onUpdateState(value => { eventVersion.current++; setState(value); setError(''); });
    const revision = eventVersion.current;
    if (bridge) void bridge.getUpdateState().then(value => {
      if (mounted.current && generation.current === cycle && eventVersion.current === revision) setState(value);
    }).catch(() => { if (mounted.current && generation.current === cycle && eventVersion.current === revision) setError('Update information could not be loaded. Close this window and try again.'); });
    else setState({ status: 'disabled', appVersion: '', autoCheck: false, message: 'Open the installed InjOffice application to manage updates.' });
    return () => { mounted.current = false; generation.current++; unsubscribe?.(); };
  }, [bridge]);

  async function act(operation: () => Promise<DesktopUpdateState>) {
    if (pending) return;
    setPending(true); setError('');
    const revision = eventVersion.current;
    const cycle = generation.current;
    try {
      const value = await operation();
      if (mounted.current && generation.current === cycle && eventVersion.current === revision) setState(value);
    } catch (reason) {
      if (mounted.current && generation.current === cycle) setError(reason instanceof Error ? reason.message : 'The update could not be completed. Please try again.');
    } finally { if (mounted.current && generation.current === cycle) setPending(false); }
  }

  const status = state?.status;
  const busy = pending || status === 'checking' || status === 'downloading' || status === 'installing';
  const progress = Math.min(100, Math.max(0, state?.percent ?? 0));
  return <dialog ref={dialog} className="updates-dialog" aria-labelledby="updates-title" onCancel={event => { event.preventDefault(); onClose(); }}>
    <header className="updates-heading"><img src={logo} alt="" /><div><h2 id="updates-title">App updates</h2><p>InjOffice{state?.appVersion ? ` · Version ${state.appVersion}` : ''}</p></div><button autoFocus className="updates-close" aria-label="Close updates" onClick={onClose}>×</button></header>
    <div className="updates-summary" role="status" aria-live="polite"><h3>{state ? labels[state.status] : 'Loading update information…'}</h3>
      {state?.version && ['available', 'downloading', 'downloaded', 'installing'].includes(state.status) && <p>Version {state.version}</p>}
      {state?.message && <p>{state.message}</p>}
      {status === 'idle' && <p>Check for the latest improvements and fixes.</p>}
      {status === 'downloaded' && <p>Your documents will be checked before restarting. You can keep working and update later.</p>}
      {status === 'downloading' && <div className="updates-progress"><progress max="100" value={Number.isFinite(progress) ? progress : undefined} aria-label="Update download progress" /><span>{Number.isFinite(progress) ? `${Math.round(progress)}%` : 'Downloading'}</span></div>}
    </div>
    {state?.releaseNotes && <section className="updates-notes" aria-labelledby="updates-notes-title"><h3 id="updates-notes-title">What’s new</h3><p>{state.releaseNotes}</p></section>}
    {error && <p className="updates-error" role="alert">{error}</p>}
    {state && status !== 'disabled' && <label className="updates-automatic"><input type="checkbox" checked={state.autoCheck} disabled={busy} onChange={event => { const enabled = event.target.checked; if (bridge) void act(() => bridge.setAutomaticUpdates(enabled)); }} />Automatically check for updates</label>}
    <div className="updates-actions"><button onClick={onClose}>{status === 'downloaded' ? 'Later' : 'Close'}</button>
      {bridge && state && status !== 'disabled' && status !== 'downloaded' && status !== 'downloading' && status !== 'installing' && <button className="primary-button" disabled={busy} onClick={() => void act(() => status === 'available' ? bridge.downloadUpdate() : bridge.checkForUpdates())}>{status === 'available' ? 'Download update' : status === 'checking' ? 'Checking…' : 'Check for updates'}</button>}
      {bridge && status === 'downloaded' && <button className="primary-button" disabled={busy} onClick={() => void act(() => bridge.installUpdate())}>{pending ? 'Preparing restart…' : 'Restart and update'}</button>}
    </div>
  </dialog>;
}
