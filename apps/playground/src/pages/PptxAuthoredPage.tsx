import { useMemo, useState } from 'react'
import { compileDeckSpecToNativeV1 } from '@injoffice/pptx-authored'
import { DsChip, DsField, DsInput, DsSelect, DsTextarea } from '../design-system/primitives'
import '../design-system/live-tools.css'
import { makeAuthoredPresentationDemo } from '../presentationDemoFixtures'

const THEMES = ['boardroom', 'midnight', 'slate', 'terra', 'forest', 'plum'] as const

export default function PptxAuthoredPage() {
  const [title, setTitle] = useState('Northstar launch review')
  const [subtitle, setSubtitle] = useState('Q3 operating brief · September 2026')
  const [theme, setTheme] = useState<(typeof THEMES)[number]>('boardroom')
  const [slideIndex, setSlideIndex] = useState(0)

  const result = useMemo(() => compileDeckSpecToNativeV1(
    makeAuthoredPresentationDemo(title, subtitle, theme),
  ), [subtitle, theme, title])

  const nativeSlide = result.ok ? result.deck.slides[slideIndex] ?? result.deck.slides[0] : undefined
  const nativeElements = result.ok ? result.deck.slides.reduce((count, slide) => count + slide.elements.length, 0) : 0

  return (
    <section className="capability-page capability-page--pptx-authored ds" data-demo-surface="pptx-authored" aria-labelledby="pptx-authored-title">
      <header className="capability-intro ds-surf-head">
        <div>
          <p className="capability-package ds-surf-pkg">@injoffice/pptx-authored</p>
          <h2 id="pptx-authored-title">Compile an authored deck into the native contract</h2>
          <p>DeckSpec stays plain JSON. The deterministic compiler validates it and emits editable <code>pptx-native/v1</code> data without a DOM or a PowerPoint process.</p>
        </div>
        <span className="capability-runtime" role="status">Runs in this browser</span>
      </header>

      <div className="capability-workspace ds-split ds-split--wide">
        <form className="capability-controls ds-split-main" onSubmit={(event) => event.preventDefault()}>
          <DsField label="Deck title">
            <DsInput value={title} maxLength={240} onChange={(event) => setTitle(event.target.value)} />
          </DsField>
          <DsField label="Subtitle">
            <DsTextarea value={subtitle} maxLength={480} rows={4} onChange={(event) => setSubtitle(event.target.value)} />
          </DsField>
          <DsField label="Authored theme">
            <DsSelect value={theme} onChange={(event) => setTheme(event.target.value as (typeof THEMES)[number])}>
              {THEMES.map((value) => <option key={value} value={value}>{value}</option>)}
            </DsSelect>
          </DsField>
          <DsField label="Inspect compiled slide">
            <DsSelect value={slideIndex} onChange={(event) => setSlideIndex(Number(event.target.value))}>
              {result.ok ? result.deck.slides.map((slide, index) => <option key={slide.id} value={index}>Slide {index + 1} · {slide.elements.length} elements</option>) : null}
            </DsSelect>
          </DsField>
          <p className="capability-note ds-muted">This credible four-slide operating review includes title, evidence, two-column plan, and decision layouts with transitions. Change any field to compile a new immutable native deck. Writing real <code>.pptx</code> bytes remains the Go patcher’s job.</p>
        </form>

        <div className="capability-result ds-split-side" aria-live="polite">
          {result.ok ? (
            <>
              <div className="capability-result-heading">
                <DsChip tone="green"><span className="capability-status capability-status--success">Contract accepted</span></DsChip>
                <strong>{result.deck.contractVersion}</strong>
              </div>
              <dl className="capability-facts ds-proof">
                <div><dt>Document</dt><dd>{result.deck.documentId}</dd></div>
                <div><dt>Origin</dt><dd>{result.deck.origin}</dd></div>
                <div><dt>Slides</dt><dd>{result.deck.slides.length}</dd></div>
                <div><dt>Native elements</dt><dd>{nativeElements}</dd></div>
                <div><dt>Theme</dt><dd>{theme}</dd></div>
                <div><dt>Compatibility</dt><dd>{result.deck.compatibility.status}</dd></div>
              </dl>
              <p className="ds-code">{nativeSlide?.elements.map((element) => element.kind).join(' · ') || 'No visual elements'}</p>
              <div className="capability-code-preview">
                <div className="capability-code-title">Slide {slideIndex + 1} native element kinds</div>
                <code>{nativeSlide?.elements.map((element) => element.kind).join(' · ') || 'No visual elements'}</code>
              </div>
            </>
          ) : (
            <>
              <DsChip tone="refuse"><span className="capability-status capability-status--refused">Compilation refused</span></DsChip>
              <ul className="capability-issues">
                {result.issues.slice(0, 8).map((issue) => (
                  <li key={`${issue.path}-${issue.code}`}><code>{issue.path}</code> {issue.message}</li>
                ))}
              </ul>
            </>
          )}
        </div>
      </div>
    </section>
  )
}
