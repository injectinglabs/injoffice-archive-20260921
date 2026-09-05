import { useMemo, useState } from 'react'
import { TARGET_FUNCTIONS } from '@injoffice/formulas'
import { createPlaygroundCalculator, playgroundCalculationRequest } from '../formulaCalculation'
import type { CalculationJobSnapshot, CalculationResult } from '../../../../packages/formulas/src/calculation'

const DECLARED_FUNCTIONS = 535

function family(name: string): string {
  if (/^(SUM|PRODUCT|ABS|ROUND|INT|MOD|POWER|SQRT|EXP|LN|LOG|CEILING|FLOOR|MROUND|TRUNC|SIGN|RAND|PI|SUBTOTAL|AGGREGATE)/.test(name)) return 'Math'
  if (/^(AVERAGE|COUNT|MAX|MIN|MEDIAN|MODE|STDEV|VAR|LARGE|SMALL|RANK|PERCENTILE|QUARTILE|CORREL|FORECAST|SLOPE|INTERCEPT|TREND|GEOMEAN)/.test(name)) return 'Statistics'
  if (/^(IF|AND|OR|NOT|XOR|SWITCH|LET|LAMBDA)/.test(name)) return 'Logic'
  if (/^(VLOOKUP|HLOOKUP|XLOOKUP|LOOKUP|INDEX|MATCH|XMATCH|OFFSET|INDIRECT|CHOOSE|ROW|COLUMN|UNIQUE|SORT|FILTER|SEQUENCE|TRANSPOSE)/.test(name)) return 'Lookup & arrays'
  if (/^(CONCAT|TEXT|LEFT|RIGHT|MID|LEN|LOWER|UPPER|PROPER|TRIM|SUBSTITUTE|REPLACE|FIND|SEARCH)/.test(name)) return 'Text'
  return 'Business functions'
}

export default function FormulasPage() {
  const [query, setQuery] = useState('')
  const [selectedName, setSelectedName] = useState('XLOOKUP')
  const matches = useMemo(() => {
    const needle = query.trim().toLowerCase()
    return TARGET_FUNCTIONS.filter((item) => !needle || item.name.toLowerCase().includes(needle) || item.formula.toLowerCase().includes(needle))
  }, [query])
  const selected = TARGET_FUNCTIONS.find((item) => item.name === selectedName) ?? matches[0] ?? TARGET_FUNCTIONS[0]!
  const calculator = useMemo(() => createPlaygroundCalculator(), [])
  const [jobFormula, setJobFormula] = useState('=2+3')
  const [job, setJob] = useState<CalculationJobSnapshot | null>(null)
  const [jobResult, setJobResult] = useState<CalculationResult | null>(null)
  const [jobError, setJobError] = useState<string | null>(null)
  const [jobCount, setJobCount] = useState(0)

  return (
    <section className="tool-page" data-demo-surface="formulas" aria-label="Formula compatibility workbench">
      <div className="tool-page__controls" role="search" aria-label="Formula search">
        <label className="tool-field tool-field--search">Search audited functions<input type="search" value={query} onChange={(event) => setQuery(event.target.value)} placeholder="XLOOKUP, date, text…" /></label>
        <span className="tool-page__status" role="status" aria-live="polite">{matches.length} matching · real-engine audit, not a name-list comparison</span>
      </div>

      <div className="tool-kpis" aria-label="Formula coverage summary">
        <div><strong>{DECLARED_FUNCTIONS}</strong><span>functions declared by the engine</span></div>
        <div><strong>{TARGET_FUNCTIONS.length}/{TARGET_FUNCTIONS.length}</strong><span>high-value formulas computing in CI</span></div>
        <div><strong>0</strong><span>#NAME? regressions in the audited set</span></div>
      </div>

      <div className="tool-page__grid">
        <section className="tool-card" aria-labelledby="formula-list-title">
          <div className="tool-card__heading"><div><span className="tool-eyebrow">Compatibility inventory</span><h2 id="formula-list-title">Audited formulas</h2></div><span className="tool-chip">{matches.length}</span></div>
          <div className="formula-list">
            {matches.map((item) => (
              <button key={item.name} type="button" aria-pressed={selected.name === item.name} className={selected.name === item.name ? 'formula-list__item formula-list__item--selected' : 'formula-list__item'} onClick={() => setSelectedName(item.name)}>
                <strong>{item.name}</strong><span>{family(item.name)}</span>
              </button>
            ))}
            {matches.length === 0 && <p className="tool-empty">No audited function matches that search.</p>}
          </div>
        </section>

        <section className="tool-card tool-card--hero" aria-labelledby="formula-detail-title">
          <div className="tool-card__heading"><div><span className="tool-eyebrow">Valid fixture invocation</span><h2 id="formula-detail-title">{selected.name}</h2></div><span className="tool-chip">{family(selected.name)}</span></div>
          <div className="formula-example"><code>{selected.formula}</code></div>
          <div className="formula-proof">
            <span className="formula-proof__mark" aria-hidden="true">✓</span>
            <div><strong>Computes in the pinned Univer engine</strong><p>CI evaluates this invocation against a real workbook and fails on name, argument, or computation regressions.</p></div>
          </div>
          <p className="tool-note">The package publishes the typed, auditable matrix. Formula calculation remains the responsibility of the selected workbook engine.</p>
          <div className="formula-example">
            <label className="tool-field">Calculation job<input value={jobFormula} onChange={(event) => setJobFormula(event.target.value)} aria-label="Formula for calculation job" /></label>
            <button className="workbench-button" type="button" onClick={() => {
              setJobError(null)
              const next = jobCount + 1
              setJobCount(next)
              const submitted = calculator.manager.submit(playgroundCalculationRequest(jobFormula, `job-${next}`))
              setJob(submitted.getState())
              void submitted.result.then((result) => {
                setJob(calculator.manager.getJob(submitted.id) ?? submitted.getState())
                setJobResult(result)
              }).catch((reason: unknown) => {
                setJob(calculator.manager.getJob(submitted.id) ?? submitted.getState())
                setJobError(reason instanceof Error ? reason.message : String(reason))
              })
            }}>Submit job</button>
            <button className="workbench-button" type="button" onClick={() => setJobFormula(selected.formula)}>Use selected formula</button>
          </div>
          {job && <p className="tool-note" role="status">{job.id} · {job.state} · {job.requestFingerprint}</p>}
          {jobResult && <p className="tool-note">{jobResult.cells.map((cell) => cell.result.kind === 'value' ? String(cell.result.value) : cell.result.kind === 'error' ? cell.result.code : 'spill').join(', ')}</p>}
          {jobError && <p className="tool-error" role="alert">{jobError}</p>}
        </section>
      </div>
    </section>
  )
}
