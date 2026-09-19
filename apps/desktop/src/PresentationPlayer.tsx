import { useEffect, useRef, useState, type ReactNode } from 'react';
import type { NativeSlide } from '@injoffice/pptx-native';
import { navigatePresentation, presentationScale, type PresentationModeState } from './presentationMode';
import './presentation-player.css';

export default function PresentationPlayer({ initial, renderSlide, onExit }: { initial: PresentationModeState; renderSlide(slide: NativeSlide, scale: number): ReactNode; onExit(): void }) {
  const [state, setState] = useState(initial); const [size, setSize] = useState({ width: 0, height: 0 });
  const [fullscreen, setFullscreen] = useState(false); const [fullscreenError, setFullscreenError] = useState('');
  const dialog = useRef<HTMLDialogElement>(null); const surface = useRef<HTMLDivElement>(null); const stage = useRef<HTMLDivElement>(null);
  const latest = useRef(onExit); latest.current = onExit; const hadFullscreen = useRef(false); const exiting = useRef(false);
  function exit() { if (exiting.current) return; exiting.current = true; latest.current(); }
  useEffect(() => {
    const modal = dialog.current!; const root = surface.current!; const previous = document.activeElement as HTMLElement | null;
    modal.showModal(); root.focus();
    const observer = new ResizeObserver(entries => { const box = entries[0]?.contentRect; if (box) setSize({ width: box.width, height: box.height }); });
    observer.observe(stage.current!);
    const changed = () => {
      const active = document.fullscreenElement === root; setFullscreen(active);
      if (!active && hadFullscreen.current) exit();
      hadFullscreen.current = active;
    };
    document.addEventListener('fullscreenchange', changed);
    return () => {
      observer.disconnect(); document.removeEventListener('fullscreenchange', changed);
      if (document.fullscreenElement === root) void document.exitFullscreen().catch(() => {});
      modal.close(); previous?.focus();
    };
  }, []);
  const slide = state.deck.slides[state.index]!;
  const scale = presentationScale(size.width, size.height, state.deck.size.cx / 9525, state.deck.size.cy / 9525);
  const diagnostics = slide.compatibility.diagnostics;
  const unsupported = slide.compatibility.status !== 'editable';
  function navigate(key: string) { const next = navigatePresentation(state, key); if (next) setState(next); else exit(); }
  async function enterFullscreen() {
    if (!surface.current?.requestFullscreen) { setFullscreenError('Fullscreen is unavailable. You can continue presenting in this window.'); return; }
    try { setFullscreenError(''); await surface.current.requestFullscreen(); } catch { setFullscreenError('Fullscreen could not be opened. You can continue presenting in this window.'); }
  }
  return <dialog className="presentation-player-dialog" ref={dialog} aria-label="Presentation view" onCancel={event => { event.preventDefault(); exit(); }}>
    <div className="presentation-player" ref={surface} tabIndex={-1} onKeyDown={event => {
      event.stopPropagation();
      if (event.key === 'Escape') { event.preventDefault(); exit(); return; }
      if (event.ctrlKey || event.metaKey || event.altKey || event.target instanceof HTMLInputElement) return;
      if (event.key === ' ' && event.target instanceof HTMLButtonElement) return;
      if (['ArrowRight','ArrowDown','ArrowLeft','ArrowUp','Home','End','PageDown','PageUp',' '].includes(event.key)) { event.preventDefault(); navigate(event.key); }
    }}>
      <header className="presentation-player-header"><span>Presentation</span><div>{!fullscreen && <button onClick={() => void enterFullscreen()}>Fullscreen</button>}<button onClick={exit} aria-label="Exit presentation">Exit <kbd>Esc</kbd></button></div></header>
      <div className="presentation-player-stage" ref={stage}>{scale > 0 && renderSlide(slide, scale)}</div>
      <footer className="presentation-player-footer"><div className="presentation-player-note">{initial.hasDraft ? 'Showing applied changes. Pending edits remain in the editor.' : 'Local preview · text wrapping may differ in PowerPoint.'}{unsupported && <details><summary>Unsupported content is preserved</summary><p>Preserved objects can appear as placeholders. This view does not reproduce every PowerPoint feature.</p>{diagnostics.length > 0 && <ul>{diagnostics.slice(0, 8).map((item, i) => <li key={i}>{item.message}</li>)}</ul>}</details>}{fullscreenError && <span role="status">{fullscreenError}</span>}</div><nav aria-label="Presentation navigation"><button aria-label="Previous slide" disabled={state.index === 0} onClick={() => navigate('ArrowLeft')}>←</button><span role="status" aria-live="polite">{state.index + 1} / {state.deck.slides.length}</span><button aria-label="Next slide" disabled={state.index === state.deck.slides.length - 1} onClick={() => navigate('ArrowRight')}>→</button></nav></footer>
    </div>
  </dialog>;
}
