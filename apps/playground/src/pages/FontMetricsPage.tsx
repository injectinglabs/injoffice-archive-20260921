import { useMemo, useState } from 'react'
import {
  classifyNativeOfficeLineBreak,
  scaleFontUnits,
  validateTextRunInput,
  type NativeOfficeLineBreakDecision,
} from '@injoffice/font-metrics/layout'

const DECISION_COPY: Record<NativeOfficeLineBreakDecision, string> = {
  allowed: 'A native layout engine may wrap at this cluster boundary.',
  prohibited: 'The clusters must stay together at this boundary.',
  unsupported: 'The v1 contract refuses to guess this boundary.',
}

export default function FontMetricsPage() {
  const [left, setLeft] = useState('Quarterly-')
  const [right, setRight] = useState('results')
  const [fontSize, setFontSize] = useState(18)
  const decision = useMemo(() => classifyNativeOfficeLineBreak(left, right), [left, right])
  const scaled = useMemo(() => scaleFontUnits(1_850, 2_048, Math.round(fontSize * 1_000)), [fontSize])
  const run = useMemo(() => validateTextRunInput({
    version: 1,
    text: `${left}${right}`,
    fontSizeMilliPoints: Math.round(fontSize * 1_000),
    font: {
      families: ['Aptos', 'Carlito'],
      weight: 400,
      style: 'normal',
      stretch: 100,
      fallbackChainIds: ['latin.default'],
    },
    script: 'Latn',
    language: 'en-US',
    direction: 'ltr',
    features: [{ tag: 'kern', value: 1 }],
  }), [fontSize, left, right])

  return (
    <section className="capability-page capability-page--font-metrics" data-demo-surface="font-metrics" aria-labelledby="font-metrics-title">
      <header className="capability-intro">
        <div>
          <p className="capability-package">@injoffice/font-metrics</p>
          <h2 id="font-metrics-title">Inspect the native text-layout contract</h2>
          <p>The browser-safe layout entry validates exact run inputs, uses integer milli-points, and makes conservative Office line-break decisions.</p>
        </div>
        <span className="capability-runtime" role="status">Contract in browser · shaping in Node</span>
      </header>

      <div className="capability-workspace">
        <div className="capability-controls">
          <fieldset className="capability-fieldset">
            <legend>Cluster boundary</legend>
            <label className="capability-field">
              Left cluster
              <input value={left} maxLength={80} onChange={(event) => setLeft(event.target.value)} />
            </label>
            <label className="capability-field">
              Right cluster
              <input value={right} maxLength={80} onChange={(event) => setRight(event.target.value)} />
            </label>
          </fieldset>
          <label className="capability-field">
            Font size
            <span className="capability-range">
              <input type="range" min="8" max="72" step="1" value={fontSize} onChange={(event) => setFontSize(Number(event.target.value))} />
              <output>{fontSize} pt</output>
            </span>
          </label>
          <div className="capability-type-sample" style={{ fontSize: `${fontSize}px` }} aria-label={`Type sample at ${fontSize} points`}>
            <span>{left}</span><span className="capability-boundary" aria-hidden="true">¦</span><span>{right}</span>
          </div>
        </div>

        <div className="capability-result" aria-live="polite">
          <div className="capability-decision" data-decision={decision}>
            <span className={`capability-status capability-status--${decision}`}>{decision}</span>
            <p>{DECISION_COPY[decision]}</p>
          </div>
          <dl className="capability-facts">
            <div><dt>Run contract</dt><dd>{run.ok ? 'valid' : 'refused'}</dd></div>
            <div><dt>Input units</dt><dd>{fontSize * 1_000} milli-pt</dd></div>
            <div><dt>Scaled ascent</dt><dd>{scaled} milli-pt</dd></div>
            <div><dt>Direction</dt><dd>LTR / Latn / en-US</dd></div>
          </dl>
          {!run.ok ? (
            <ul className="capability-issues">
              {run.issues.slice(0, 8).map((issue, index) => <li key={`${issue.path}-${index}`}><code>{issue.path}</code> {issue.message}</li>)}
            </ul>
          ) : null}
          <div className="capability-runtime-boundary">
            <strong>Host runtime boundary</strong>
            <p>System-font discovery and canonical HarfBuzz shaping require Node.js 22, licensed or document-embedded font bytes, and host-injected resolver policy. No browser font or canvas measurement is used here.</p>
          </div>
        </div>
      </div>
    </section>
  )
}
