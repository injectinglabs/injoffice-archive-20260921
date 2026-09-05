import { useState, type CSSProperties } from 'react'
import {
  BUILTIN_THEMES,
  DeckCanvasView,
  DeckEditorPanel,
  deckFromOutline,
  resetSlideIds,
  type AnimDirection,
} from '@injoffice/slides'
import { playgroundDeckAudit, updateSlideTransition, type SlideTransitionChoice } from '../slideQc'
import { PRESENTATION_DEMO_OUTLINE, PRESENTATION_DEMO_TITLE } from '../presentationDemoFixtures'

export default function SlidesPage() {
  const [outline, setOutline] = useState(PRESENTATION_DEMO_OUTLINE)
  const [themeId, setThemeId] = useState(BUILTIN_THEMES[0]!.id)
  const [at, setAt] = useState(0)
  const [spec, setSpec] = useState(() => {
    resetSlideIds()
    const next = deckFromOutline(PRESENTATION_DEMO_TITLE, PRESENTATION_DEMO_OUTLINE)
    next.theme = BUILTIN_THEMES[0]!.id
    return next
  })
  const theme = BUILTIN_THEMES.find((t) => t.id === themeId) ?? BUILTIN_THEMES[0]!
  const audit = playgroundDeckAudit(spec)
  const transition = spec.slides[at]?.transition

  return (
    <div className="platen-fill" data-demo-surface="slides">
      <div className="toolstrip univer-toolbar" role="toolbar" aria-label="Presentation tools">
        <label className="tool-field">
          <span>Theme</span>
          <select
            value={themeId}
            aria-label="Deck theme"
            onChange={(e) => {
              const id = e.target.value
              setThemeId(id)
              setSpec((prev) => ({ ...prev, theme: id }))
            }}
          >
            {BUILTIN_THEMES.map((t) => (
              <option key={t.id} value={t.id}>{t.id}</option>
            ))}
          </select>
        </label>
        <label className="tool-field">
          <span>Transition</span>
          <select
            value={transition?.kind ?? 'none'}
            aria-label="Selected slide transition"
            onChange={(event) => setSpec((current) => updateSlideTransition(current, at, event.target.value as SlideTransitionChoice, transition?.direction))}
          >
            <option value="none">None</option>
            <option value="fade">Fade</option>
            <option value="push">Push</option>
            <option value="wipe">Wipe</option>
          </select>
        </label>
        {transition && transition.kind !== 'fade' ? (
          <label className="tool-field">
            <span>Direction</span>
            <select
              value={transition.direction ?? 'left'}
              aria-label="Selected slide transition direction"
              onChange={(event) => setSpec((current) => updateSlideTransition(current, at, transition.kind, event.target.value as AnimDirection))}
            >
              <option value="left">Left</option>
              <option value="right">Right</option>
              <option value="top">Top</option>
              <option value="bottom">Bottom</option>
            </select>
          </label>
        ) : null}
        <button
          type="button"
          className="workbench-button workbench-button--primary"
          onClick={() => {
            resetSlideIds()
            const next = deckFromOutline(PRESENTATION_DEMO_TITLE, outline)
            next.theme = themeId
            setSpec(next)
            setAt(0)
          }}
        >
          Build slides
        </button>
        <span className="univer-toolbar__status">
          {spec.slides.length} {spec.slides.length === 1 ? 'slide' : 'slides'} · built from DeckSpec; no PPTX file is loaded
        </span>
      </div>
      <div className="split">
        <div
          className="split-main slides-workspace"
          style={{ '--slides-workspace-background': theme.background } as CSSProperties}
        >
          <DeckCanvasView spec={spec} at={at} onAtChange={setAt} width={720} editable />
          <textarea
            className="slides-outline"
            aria-label="Deck outline"
            value={outline}
            onChange={(e) => setOutline(e.target.value)}
          />
        </div>
        <aside className="split-side">
          <DeckEditorPanel spec={spec} at={at} onChange={setSpec} onSelect={setAt} />
          <div className="ioc-panel">
            <strong>Layout QC</strong>
            <p>{audit.issues.length === 0 ? 'No estimated overflow or overlap.' : `${audit.issues.length} issue${audit.issues.length === 1 ? '' : 's'}.`}</p>
            <pre>{audit.report}</pre>
          </div>
        </aside>
      </div>
    </div>
  )
}
