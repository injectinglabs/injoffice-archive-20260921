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
import { DsCallout, DsChip, DsField, DsInput, DsSelect } from '../design-system/primitives'
import '../design-system/live-tools.css'
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

function fidelityTone(fidelity: ShapePreviewFidelity): 'green' | 'plain' | 'refuse' {
  if (fidelity === 'distinct') return 'green'
  if (fidelity === 'generic') return 'refuse'
  return 'plain'
}

export default function ShapesPage() {
  const [category, setCategory] = useState(SHAPE_CATEGORIES[0]!.label)
  const [query, setQuery] = useState('')
  const [selected, setSelected] = useState<ShapeKind>('roundRect')
  const [fill, setFill] = useState('#e7f4f8')
  const [stroke, setStroke] = useState('#0f6f8f')
  const kinds = useMemo(() => {
    const group = SHAPE_CATEGORIES.find((candidate) => candidate.label === category)
    return (group?.kinds ?? []).filter((kind) => shapeLabel(kind).toLowerCase().includes(query.toLowerCase())) as ShapeKind[]
  }, [category, query])
  const defaults = SHAPE_DEFAULTS[selected]
  const wire = useMemo(() => playgroundShapeWire(selected, fill, stroke), [fill, selected, stroke])
  const selectedFidelity = shapePreviewFidelity(selected)

  return (
    <div className="ds">
    <section className="tool-page" data-demo-surface="shapes" aria-label="Shape geometry workbench">
      <div className="tool-page__controls ds-workstrip" role="toolbar" aria-label="Shape controls">
        <DsField label="Category">
          <DsSelect value={category} onChange={(event) => setCategory(event.target.value)}>
            {SHAPE_CATEGORIES.map((item) => <option key={item.label}>{item.label}</option>)}
          </DsSelect>
        </DsField>
        <DsField label="Find a shape">
          <DsInput type="search" value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Decision, star, callout…" />
        </DsField>
        <span className="tool-page__status ds-muted">121 OOXML-native shape kinds</span>
      </div>

      <div className="tool-page__grid ds-split">
        <section className="tool-card ds-split-main" aria-labelledby="shape-library-title">
          <span className="ds-eyebrow tool-eyebrow">Preset library</span>
          <div className="ds-row" style={{ marginBottom: 8 }}>
            <h2 id="shape-library-title">{category}</h2>
            <DsChip>{kinds.length} shown</DsChip>
          </div>
          <div className="shape-library ds-shape-lib">
            {kinds.map((kind) => {
              const fidelity = shapePreviewFidelity(kind)
              return (
                <button key={kind} type="button" aria-pressed={selected === kind} className={selected === kind ? 'shape-library__item shape-library__item--selected' : 'shape-library__item'} onClick={() => setSelected(kind)}>
                  <ShapePreview kind={kind} fill="#e7f4f8" stroke="#0f6f8f" strokeWidth={2} decorative />
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

        <section className="tool-card tool-card--hero ds-split-side" aria-labelledby="shape-inspector-title">
          <span className="ds-eyebrow tool-eyebrow">Browser preview fidelity</span>
          <h2 id="shape-inspector-title">{shapeLabel(selected)}</h2>
          <div className="shape-stage ds-shape-stage">
            <ShapePreview kind={selected} fill={fill} stroke={stroke} strokeWidth={2} ariaLabel={`${shapeLabel(selected)} browser preview`} />
          </div>
          <DsChip tone={fidelityTone(selectedFidelity)}>{FIDELITY_LABEL[selectedFidelity]}</DsChip>
          <p className={`shape-preview-note shape-preview-note--${selectedFidelity}`} role="status">
            <strong>{FIDELITY_LABEL[selectedFidelity]}.</strong> {fidelityNote(selectedFidelity)}
          </p>
          {selectedFidelity !== 'distinct' && (
            <DsCallout tone={selectedFidelity === 'generic' ? 'refuse' : 'note'} title={selectedFidelity === 'generic' ? 'Generic placeholder' : 'Approximate geometry'}>
              {fidelityNote(selectedFidelity)}
            </DsCallout>
          )}
          <div className="shape-controls ds-field-row">
            <DsField label="Fill"><input type="color" value={fill} onChange={(event) => setFill(event.target.value)} /></DsField>
            <DsField label="Stroke"><input type="color" value={stroke} onChange={(event) => setStroke(event.target.value)} /></DsField>
          </div>
          <dl className="tool-metrics ds-proof">
            <div><dt>Default fill</dt><dd>{defaults.fill || 'transparent'}</dd></div>
            <div><dt>Default stroke</dt><dd>{defaults.stroke || 'none'}</dd></div>
            <div><dt>OOXML preset</dt><dd>{selected}</dd></div>
            <div><dt>Native wire</dt><dd>{wire.shapes[0] ? `${wire.shapes[0].sheetName} ${wire.shapes[0].kind}` : wire.skipped[0]}</dd></div>
          </dl>
          <pre className="ds-code tool-note">{JSON.stringify(wire.shapes[0] ?? { skipped: wire.skipped }, null, 2)}</pre>
        </section>
      </div>
    </section>
    </div>
  )
}
