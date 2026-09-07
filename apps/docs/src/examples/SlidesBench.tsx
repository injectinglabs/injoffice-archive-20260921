import { useMemo, useState } from 'react'
import { compileDeckSpecToNativeV1 } from '@injoffice/pptx-authored'
import { addSlide, updateSlide, type DeckSpec } from '@injoffice/slides'
import { LiveBench } from '../components/LiveBench'

const START: DeckSpec = {
  id: 'docs-deck',
  title: 'InjOffice',
  theme: 'slate',
  slides: [
    { id: 's1', kind: 'title', title: 'Native Office files', subtitle: 'Original bytes stay authoritative' },
    { id: 's2', kind: 'bullets', title: 'Pipeline', bullets: ['Extract', 'Mutate', 'Apply', 'Reopen'] },
  ],
}

export function SlidesBench() {
  const [spec, setSpec] = useState(START)
  const compiled = useMemo(() => compileDeckSpecToNativeV1(spec), [spec])
  const title = spec.slides[0]?.title ?? ''
  const nativeElements = compiled.ok
    ? compiled.deck.slides.reduce((count, slide) => count + slide.elements.length, 0)
    : 0
  return (
    <LiveBench title="Live example" hint="@injoffice/slides + @injoffice/pptx-authored · compileDeckSpecToNativeV1">
      <div className="bench-controls">
        <label className="field" style={{ flex: 1 }}>
          Title slide
          <input value={title} onChange={(event) => setSpec((current) => updateSlide(current, 0, { title: event.target.value }))} />
        </label>
        <button type="button" className="bench-button" onClick={() => setSpec((current) => addSlide(current, current.slides.length - 1, 'bullets'))}>Add slide</button>
      </div>
      {compiled.ok ? (
        <div>
          <p><span className="badge badge--pass">{compiled.deck.contractVersion}</span> {compiled.deck.slides.length} native slide{compiled.deck.slides.length === 1 ? '' : 's'} · {nativeElements} elements</p>
          <table className="grid-table">
            <thead><tr><th>Slide</th><th>Native kinds</th></tr></thead>
            <tbody>
              {compiled.deck.slides.map((slide, index) => (
                <tr key={slide.id}>
                  <td>{index + 1}. {spec.slides[index]?.title || slide.id}</td>
                  <td>{slide.elements.map((element) => element.kind).join(' · ') || 'none'}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : (
        <div>
          <p><span className="badge badge--patch">refused</span></p>
          <ul>{compiled.issues.map((issue) => <li key={issue.path}><code>{issue.code}</code> {issue.message}</li>)}</ul>
        </div>
      )}
    </LiveBench>
  )
}
