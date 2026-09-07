import { useMemo, useState } from 'react'
import { classifyNativeOfficeLineBreak, validateTextRunInput } from '@injoffice/font-metrics/layout'
import { LiveBench } from '../components/LiveBench'

export function FontMetricsBench() {
  const [left, setLeft] = useState('Office')
  const [right, setRight] = useState(' file')
  const decision = useMemo(() => classifyNativeOfficeLineBreak(left, right), [left, right])
  const validated = useMemo(() => validateTextRunInput({
    version: 1,
    text: `${left}${right}`,
    fontSizeMilliPoints: 18_000,
    font: { families: ['Calibri'], weight: 400, style: 'normal', stretch: 100 },
    script: 'latn',
    language: 'en',
    direction: 'ltr',
  }), [left, right])
  return (
    <LiveBench title="Live example" hint="@injoffice/font-metrics/layout · classifyNativeOfficeLineBreak">
      <div className="bench-controls">
        <label className="field" style={{ flex: 1 }}>
          Left cluster
          <input value={left} onChange={(event) => setLeft(event.target.value)} spellCheck={false} />
        </label>
        <label className="field" style={{ flex: 1 }}>
          Right cluster
          <input value={right} onChange={(event) => setRight(event.target.value)} spellCheck={false} />
        </label>
      </div>
      <p>Line break at the boundary: <span className={`badge ${decision === 'unsupported' ? 'badge--patch' : 'badge--pass'}`}>{decision}</span></p>
      {validated.ok
        ? <p>Text-run contract: <span className="badge badge--pass">valid</span> {validated.value.text.length} UTF-16 units</p>
        : (
          <div>
            <p>Text-run contract: <span className="badge badge--patch">refused</span></p>
            <ul>{validated.issues.map((issue) => <li key={issue.path}><code>{issue.code}</code> {issue.path} {issue.message}</li>)}</ul>
          </div>
        )}
    </LiveBench>
  )
}
