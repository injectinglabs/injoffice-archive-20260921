import { useEffect, useMemo, useState } from 'react'
import { csvToGrid, jsonToGrid } from '@injoffice/connectors'
import { preprocessConnectorGrid } from '../connectorPreprocess'

const JSON_SAMPLE = JSON.stringify({
  data: {
    accounts: [
      { account: 'Northwind', owner: 'Maya', health: 'Healthy', arr: 128000 },
      { account: 'Contoso', owner: 'Theo', health: 'Watch', arr: 94000 },
      { account: 'Globex', owner: 'Inez', health: 'Healthy', arr: 151000 },
    ],
  },
}, null, 2)

const CSV_SAMPLE = 'account,owner,health,arr\nNorthwind,Maya,Healthy,128000\nContoso,Theo,Watch,94000\nGlobex,Inez,Healthy,151000'

export default function ConnectorsPage() {
  const [format, setFormat] = useState<'json' | 'csv'>('json')
  const [path, setPath] = useState('data.accounts')
  const [source, setSource] = useState(JSON_SAMPLE)

  const result = useMemo(() => {
    try {
      return { grid: format === 'json' ? jsonToGrid(JSON.parse(source), path) : csvToGrid(source), error: '' }
    } catch (reason) {
      return { grid: [] as (string | number | null)[][], error: reason instanceof Error ? reason.message : String(reason) }
    }
  }, [format, path, source])

  const [processed, setProcessed] = useState<{ grid: (string | number | boolean | null)[][]; fingerprint: string; stages: string[] } | null>(null)
  const [processError, setProcessError] = useState<string | null>(null)

  useEffect(() => {
    if (result.error || result.grid.length === 0) {
      setProcessed(null)
      return
    }
    let cancelled = false
    void preprocessConnectorGrid(result.grid).then((next) => {
      if (cancelled) return
      setProcessError(null)
      setProcessed({ grid: next.grid, fingerprint: next.fingerprint, stages: next.stages.map((stage) => stage.id) })
    }).catch((reason: unknown) => {
      if (cancelled) return
      setProcessed(null)
      setProcessError(reason instanceof Error ? reason.message : String(reason))
    })
    return () => { cancelled = true }
  }, [result])

  const chooseFormat = (next: 'json' | 'csv') => {
    setFormat(next)
    setSource(next === 'json' ? JSON_SAMPLE : CSV_SAMPLE)
    setPath(next === 'json' ? 'data.accounts' : '')
  }

  return (
    <section className="tool-page" data-demo-surface="connectors" aria-label="Data connector workbench">
      <div className="tool-page__controls" role="toolbar" aria-label="Connector controls">
        <span className="tool-segment" role="group" aria-label="Source format">
          <button type="button" aria-pressed={format === 'json'} onClick={() => chooseFormat('json')}>JSON</button>
          <button type="button" aria-pressed={format === 'csv'} onClick={() => chooseFormat('csv')}>CSV</button>
        </span>
        {format === 'json' && <label className="tool-field">Dot path<input value={path} onChange={(event) => setPath(event.target.value)} placeholder="data.items" /></label>}
        <button className="workbench-button" type="button" onClick={() => chooseFormat(format)}>Reset sample</button>
        <span className="tool-page__status">Host-neutral · credentials never enter the package</span>
      </div>

      <div className="tool-page__grid">
        <section className="tool-card" aria-labelledby="connector-source-title">
          <div className="tool-card__heading"><div><span className="tool-eyebrow">Host-supplied payload</span><h2 id="connector-source-title">{format.toUpperCase()} source</h2></div><span className="tool-chip">updates as you type</span></div>
          <textarea className="tool-editor" value={source} onChange={(event) => setSource(event.target.value)} spellCheck={false} aria-label={`${format.toUpperCase()} source`} />
          {result.error && <p className="tool-error" role="alert">{result.error}</p>}
        </section>
        <section className="tool-card tool-card--hero" aria-labelledby="connector-grid-title">
          <div className="tool-card__heading"><div><span className="tool-eyebrow">Normalized bound range</span><h2 id="connector-grid-title">Spreadsheet-ready grid</h2></div><span className="tool-chip">{Math.max(0, result.grid.length - 1)} rows</span></div>
          <div className="tool-table-wrap">
            <table className="tool-table">
              <caption className="visually-hidden">Normalized connector data</caption>
              <thead><tr>{(result.grid[0] ?? []).map((cell, index) => <th key={index}>{String(cell ?? '')}</th>)}</tr></thead>
              <tbody>{result.grid.slice(1).map((row, rowIndex) => <tr key={rowIndex}>{row.map((cell, index) => <td key={index}>{typeof cell === 'number' ? cell.toLocaleString() : String(cell ?? '')}</td>)}</tr>)}</tbody>
            </table>
          </div>
          <p className="tool-note">The same normalization rules back bound-range refreshes. Authentication, URL policy, and secret storage stay with the embedding host.</p>
        </section>
        <section className="tool-card" aria-labelledby="connector-preprocess-title">
          <div className="tool-card__heading"><div><span className="tool-eyebrow">Deterministic pipeline</span><h2 id="connector-preprocess-title">Range preprocess</h2></div><span className="tool-chip">{processed?.stages.length ?? 0} stages</span></div>
          {processError && <p className="tool-error" role="alert">{processError}</p>}
          {processed && (
            <>
              <p className="tool-note">{processed.stages.join(' → ')} · {processed.fingerprint}</p>
              <div className="tool-table-wrap">
                <table className="tool-table">
                  <caption className="visually-hidden">Preprocessed connector data</caption>
                  <thead><tr>{(processed.grid[0] ?? []).map((cell, index) => <th key={index}>{String(cell ?? '')}</th>)}</tr></thead>
                  <tbody>{processed.grid.slice(1).map((row, rowIndex) => <tr key={rowIndex}>{row.map((cell, index) => <td key={index}>{typeof cell === 'number' ? cell.toLocaleString() : String(cell ?? '')}</td>)}</tr>)}</tbody>
                </table>
              </div>
            </>
          )}
        </section>
      </div>
    </section>
  )
}
