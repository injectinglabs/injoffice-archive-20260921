import { useMemo, useState } from 'react'
import {
  SHAPE_CATEGORIES,
  SHAPE_DEFAULTS,
  ShapePreview,
  shapePreviewFidelity,
  shapeLabel,
  type ShapeKind,
  type ShapePreviewFidelity,
} from '@injoffice/shapes'
import { playgroundShapeWire } from '../shapeWire'

const FIDELITY_LABEL: Record<ShapePreviewFidelity, string> = {
  distinct: 'Distinct browser preview',
  approximated: 'Approximated browser preview',
  generic: 'Generic browser preview',
}

function fidelityNote(fidelity: ShapePreviewFidelity) {
  if (fidelity === 'distinct') return 'The browser renderer has distinct geometry for this preset.'
  if (fidelity === 'approximated') return 'The browser preview simplifies this preset. Saved Office files keep its exact OOXML preset name.'
  return 'A labeled placeholder is shown for this long-tail preset. Saved Office files keep its exact OOXML preset name.'
}

export default function ShapesPage() {
  const [category, setCategory] = useState(SHAPE_CATEGORIES[0]!.label)
  const [query, setQuery] = useState('')
  const [selected, setSelected] = useState<ShapeKind>('roundRect')
  const [fill, setFill] = useState('#dff7ef')
  const [stroke, setStroke] = useState('#087f6b')
  const kinds = useMemo(() => {
    const group = SHAPE_CATEGORIES.find((candidate) => candidate.label === category)
    return (group?.kinds ?? []).filter((kind) => shapeLabel(kind).toLowerCase().includes(query.toLowerCase())) as ShapeKind[]
  }, [category, query])
  const defaults = SHAPE_DEFAULTS[selected]
  const wire = useMemo(() => playgroundShapeWire(selected, fill, stroke), [fill, selected, stroke])
  const selectedFidelity = shapePreviewFidelity(selected)

  return (
    <section className="tool-page" data-demo-surface="shapes" aria-label="Shape geometry workbench">
      <div className="tool-page__controls" role="toolbar" aria-label="Shape controls">
        <label className="tool-field">Category<select value={category} onChange={(event) => setCategory(event.target.value)}>{SHAPE_CATEGORIES.map((item) => <option key={item.label}>{item.label}</option>)}</select></label>
        <label className="tool-field">Find a shape<input type="search" value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Decision, star, callout…" /></label>
        <span className="tool-page__status">121 OOXML-native shape kinds</span>
      </div>

      <div className="tool-page__grid">
        <section className="tool-card" aria-labelledby="shape-library-title">
          <div className="tool-card__heading"><div><span className="tool-eyebrow">Preset library</span><h2 id="shape-library-title">{category}</h2></div><span className="tool-chip">{kinds.length} shown</span></div>
          <div className="shape-library">
            {kinds.map((kind) => {
              const fidelity = shapePreviewFidelity(kind)
              return (
                <button key={kind} type="button" aria-pressed={selected === kind} className={selected === kind ? 'shape-library__item shape-library__item--selected' : 'shape-library__item'} onClick={() => setSelected(kind)}>
                  <ShapePreview kind={kind} fill="#d8e3f6" stroke="#62718d" strokeWidth={2} decorative />
                  <span className="shape-library__name">{shapeLabel(kind)}</span>
                  {fidelity !== 'distinct' && (
                    <span className={`shape-library__fidelity shape-library__fidelity--${fidelity}`}>
                      {fidelity === 'generic' ? 'generic' : 'approx.'}
                    </span>
                  )}
                </button>
              )
            })}
          </div>
        </section>

        <section className="tool-card tool-card--hero" aria-labelledby="shape-inspector-title">
          <div className="tool-card__heading"><div><span className="tool-eyebrow">Browser preview fidelity</span><h2 id="shape-inspector-title">{shapeLabel(selected)}</h2></div><span className="tool-chip">{selected}</span></div>
          <div className="shape-stage">
            <ShapePreview kind={selected} fill={fill} stroke={stroke} strokeWidth={2} ariaLabel={`${shapeLabel(selected)} browser preview`} />
          </div>
          <p className={`shape-preview-note shape-preview-note--${selectedFidelity}`} role="status">
            <strong>{FIDELITY_LABEL[selectedFidelity]}.</strong> {fidelityNote(selectedFidelity)}
          </p>
          <div className="shape-controls">
            <label className="tool-field">Fill<input type="color" value={fill} onChange={(event) => setFill(event.target.value)} /></label>
            <label className="tool-field">Stroke<input type="color" value={stroke} onChange={(event) => setStroke(event.target.value)} /></label>
          </div>
          <dl className="tool-metrics">
            <div><dt>Default fill</dt><dd>{defaults.fill || 'transparent'}</dd></div>
            <div><dt>Default stroke</dt><dd>{defaults.stroke || 'none'}</dd></div>
            <div><dt>OOXML preset</dt><dd>{selected}</dd></div>
            <div><dt>Native wire</dt><dd>{wire.shapes[0] ? `${wire.shapes[0].sheetName} ${wire.shapes[0].kind}` : wire.skipped[0]}</dd></div>
          </dl>
          <pre className="tool-note">{JSON.stringify(wire.shapes[0] ?? { skipped: wire.skipped }, null, 2)}</pre>
        </section>
      </div>
    </section>
  )
}
