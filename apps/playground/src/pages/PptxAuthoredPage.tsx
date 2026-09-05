import { useMemo, useState } from 'react'
import { compileDeckSpecToNativeV1 } from '@injoffice/pptx-authored'
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
    <section className="capability-page capability-page--pptx-authored" data-demo-surface="pptx-authored" aria-labelledby="pptx-authored-title">
      <header className="capability-intro">
        <div>
          <p className="capability-package">@injoffice/pptx-authored</p>
          <h2 id="pptx-authored-title">Compile an authored deck into the native contract</h2>
          <p>DeckSpec stays plain JSON. The deterministic compiler validates it and emits editable <code>pptx-native/v1</code> data without a DOM or a PowerPoint process.</p>
        </div>
        <span className="capability-runtime" role="status">Runs in this browser</span>
      </header>

      <div className="capability-workspace">
        <form className="capability-controls" onSubmit={(event) => event.preventDefault()}>
          <label className="capability-field">
            Deck title
            <input value={title} maxLength={240} onChange={(event) => setTitle(event.target.value)} />
          </label>
          <label className="capability-field">
            Subtitle
            <textarea value={subtitle} maxLength={480} rows={4} onChange={(event) => setSubtitle(event.target.value)} />
          </label>
          <label className="capability-field">
            Authored theme
            <select value={theme} onChange={(event) => setTheme(event.target.value as (typeof THEMES)[number])}>
              {THEMES.map((value) => <option key={value} value={value}>{value}</option>)}
            </select>
          </label>
          <label className="capability-field">
            Inspect compiled slide
            <select value={slideIndex} onChange={(event) => setSlideIndex(Number(event.target.value))}>
              {result.ok ? result.deck.slides.map((slide, index) => <option key={slide.id} value={index}>Slide {index + 1} · {slide.elements.length} elements</option>) : null}
            </select>
          </label>
          <p className="capability-note">This credible four-slide operating review includes title, evidence, two-column plan, and decision layouts with transitions. Change any field to compile a new immutable native deck. Writing real <code>.pptx</code> bytes remains the Go patcher’s job.</p>
        </form>

        <div className="capability-result" aria-live="polite">
          {result.ok ? (
            <>
              <div className="capability-result-heading">
                <span className="capability-status capability-status--success">Contract accepted</span>
                <strong>{result.deck.contractVersion}</strong>
              </div>
              <dl className="capability-facts">
                <div><dt>Document</dt><dd>{result.deck.documentId}</dd></div>
                <div><dt>Origin</dt><dd>{result.deck.origin}</dd></div>
                <div><dt>Slides</dt><dd>{result.deck.slides.length}</dd></div>
                <div><dt>Native elements</dt><dd>{nativeElements}</dd></div>
                <div><dt>Compatibility</dt><dd>{result.deck.compatibility.status}</dd></div>
              </dl>
              <div className="capability-code-preview">
                <div className="capability-code-title">Slide {slideIndex + 1} native element kinds</div>
                <code>{nativeSlide?.elements.map((element) => element.kind).join(' · ') || 'No visual elements'}</code>
              </div>
            </>
          ) : (
            <>
              <span className="capability-status capability-status--refused">Compilation refused</span>
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
