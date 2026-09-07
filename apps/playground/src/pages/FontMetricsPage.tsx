import { useMemo, useState } from 'react'
import {
  classifyNativeOfficeLineBreak,
  scaleFontUnits,
  validateTextRunInput,
  type NativeOfficeLineBreakDecision,
} from '@injoffice/font-metrics/layout'
import { DsCallout, DsChip, DsField, DsInput } from '../design-system/primitives'
import '../design-system/live-tools.css'

const DECISION_COPY: Record<NativeOfficeLineBreakDecision, string> = {
  allowed: 'A native layout engine may wrap at this cluster boundary.',
  prohibited: 'The clusters must stay together at this boundary.',
  unsupported: 'The v1 contract refuses to guess this boundary.',
}

function decisionTone(decision: NativeOfficeLineBreakDecision): 'green' | 'refuse' | 'plain' {
  if (decision === 'allowed') return 'green'
  if (decision === 'prohibited') return 'refuse'
  return 'plain'
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
    <section className="capability-page capability-page--font-metrics ds" data-demo-surface="font-metrics" aria-labelledby="font-metrics-title">
      <header className="capability-intro ds-surf-head">
        <div>
          <p className="capability-package ds-surf-pkg">@injoffice/font-metrics</p>
          <h2 id="font-metrics-title">Inspect the native text-layout contract</h2>
          <p>The browser-safe layout entry validates exact run inputs, uses integer milli-points, and makes conservative Office line-break decisions.</p>
        </div>
        <span className="capability-runtime" role="status">Contract in browser · shaping in Node</span>
      </header>

      <div className="capability-workspace ds-split ds-split--wide">
        <div className="capability-controls ds-split-main">
          <fieldset className="capability-fieldset ds-panel">
            <legend>Cluster boundary</legend>
            <DsField label="Left cluster">
              <DsInput value={left} maxLength={80} onChange={(event) => setLeft(event.target.value)} />
            </DsField>
            <DsField label="Right cluster">
              <DsInput value={right} maxLength={80} onChange={(event) => setRight(event.target.value)} />
            </DsField>
          </fieldset>
          <DsField label="Font size">
            <span className="capability-range range-field">
              <input type="range" min="8" max="72" step="1" value={fontSize} onChange={(event) => setFontSize(Number(event.target.value))} />
              <output>{fontSize} pt</output>
            </span>
          </DsField>
          <p className="capability-type-sample ds-type-sample" style={{ fontSize: `${fontSize}px` }} aria-label={`Type sample at ${fontSize} points`}>
            <span>{left}</span><span className="capability-boundary ds-boundary" aria-hidden="true">¦</span><span>{right}</span>
          </p>
        </div>

        <div className="capability-result ds-split-side" aria-live="polite">
          <div className="capability-decision" data-decision={decision}>
            <DsChip tone={decisionTone(decision)}>
              <span className={`capability-status capability-status--${decision}`}>{decision}</span>
            </DsChip>
            <p className="ds-muted">{DECISION_COPY[decision]}</p>
          </div>
          <dl className="capability-facts ds-proof">
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
          <DsCallout tone="note" title="Host runtime boundary">
            System-font discovery and canonical HarfBuzz shaping require Node.js 22, licensed or document-embedded font bytes, and host-injected resolver policy. No browser font or canvas measurement is used here.
          </DsCallout>
          <div className="capability-runtime-boundary">
            <strong>Host runtime boundary</strong>
            <p>No browser font or canvas measurement is used here.</p>
          </div>
        </div>
      </div>
    </section>
  )
}
