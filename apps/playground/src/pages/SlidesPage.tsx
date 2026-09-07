import { useState, type CSSProperties } from 'react'
import {
  BUILTIN_THEMES,
  DeckCanvasView,
  DeckEditorPanel,
  deckFromOutline,
  resetSlideIds,
  type AnimDirection,
} from '@injoffice/slides'
import { DsButton, DsField, DsSelect } from '../design-system/primitives'
import '../design-system/live-create-edit.css'
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
    <div className="platen-fill ds" data-demo-surface="slides">
      <div className="toolstrip univer-toolbar ds-workstrip" role="toolbar" aria-label="Presentation tools">
        <DsField label="Theme">
          <DsSelect
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
          </DsSelect>
        </DsField>
        <DsField label="Transition">
          <DsSelect
            value={transition?.kind ?? 'none'}
            aria-label="Selected slide transition"
            onChange={(event) => setSpec((current) => updateSlideTransition(current, at, event.target.value as SlideTransitionChoice, transition?.direction))}
          >
            <option value="none">None</option>
            <option value="fade">Fade</option>
            <option value="push">Push</option>
            <option value="wipe">Wipe</option>
          </DsSelect>
        </DsField>
        {transition && transition.kind !== 'fade' ? (
          <DsField label="Direction">
            <DsSelect
              value={transition.direction ?? 'left'}
              aria-label="Selected slide transition direction"
              onChange={(event) => setSpec((current) => updateSlideTransition(current, at, transition.kind, event.target.value as AnimDirection))}
            >
              <option value="left">Left</option>
              <option value="right">Right</option>
              <option value="top">Top</option>
              <option value="bottom">Bottom</option>
            </DsSelect>
          </DsField>
        ) : null}
        <DsButton
          variant="filled"
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
        </DsButton>
        <span className="univer-toolbar__status ds-muted">
          {spec.slides.length} {spec.slides.length === 1 ? 'slide' : 'slides'} · built from DeckSpec; no PPTX file is loaded
        </span>
      </div>
      <div className="split ds-split">
        <div className="split-main slides-workspace ds-split-main">
          <div
            className="ds-slide live-slide-stage"
            style={{ '--slides-workspace-background': theme.background } as CSSProperties}
          >
            <DeckCanvasView spec={spec} at={at} onAtChange={setAt} width={720} editable />
          </div>
          <textarea
            className="slides-outline ds-outline"
            aria-label="Deck outline"
            value={outline}
            onChange={(e) => setOutline(e.target.value)}
          />
        </div>
        <aside className="split-side ds-split-side">
          {spec.slides.map((slide, index) => (
            <button
              type="button"
              key={slide.id}
              className="ds-pick"
              aria-pressed={at === index}
              onClick={() => setAt(index)}
            >
              <strong>{index + 1}. {slide.title || slide.quote || slide.kind}</strong>
            </button>
          ))}
          <DeckEditorPanel spec={spec} at={at} onChange={setSpec} onSelect={setAt} />
          <div className="ioc-panel ds-panel">
            <span className="ds-eyebrow">Layout QC</span>
            <p className="ds-muted">{audit.issues.length === 0 ? 'No estimated overflow or overlap.' : `${audit.issues.length} issue${audit.issues.length === 1 ? '' : 's'}.`}</p>
            <pre className="ds-code">{audit.report}</pre>
          </div>
        </aside>
      </div>
    </div>
  )
}
