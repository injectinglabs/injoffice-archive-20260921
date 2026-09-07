import { useEffect, useMemo, useState } from 'react'
import { csvToGrid, jsonToGrid } from '@injoffice/connectors'
import { DsButton, DsChip, DsField, DsInput, DsSegment } from '../design-system/primitives'
import '../design-system/live-tools.css'
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
    <div className="ds">
    <section className="tool-page" data-demo-surface="connectors" aria-label="Data connector workbench">
      <div className="tool-page__controls ds-workstrip" role="toolbar" aria-label="Connector controls">
        <DsSegment
          label="Source format"
          value={format}
          onChange={(id) => chooseFormat(id as 'json' | 'csv')}
          options={[{ id: 'json', label: 'JSON' }, { id: 'csv', label: 'CSV' }]}
        />
        {format === 'json' && (
          <DsField label="Dot path">
            <DsInput value={path} onChange={(event) => setPath(event.target.value)} placeholder="data.items" />
          </DsField>
        )}
        <DsButton variant="outlined" className="workbench-button" onClick={() => chooseFormat(format)}>Reset sample</DsButton>
        <span className="tool-page__status ds-muted">Host-neutral · credentials never enter the package</span>
      </div>

      <div className="tool-page__grid ds-split ds-split--wide">
        <section className="tool-card ds-split-main" aria-labelledby="connector-source-title">
          <span className="ds-eyebrow tool-eyebrow">Host-supplied payload</span>
          <h2 id="connector-source-title">{format.toUpperCase()} source</h2>
          <textarea className="tool-editor ds-outline" value={source} onChange={(event) => setSource(event.target.value)} spellCheck={false} aria-label={`${format.toUpperCase()} source`} />
          {result.error && <p className="tool-error" role="alert">{result.error}</p>}
        </section>
        <section className="tool-card tool-card--hero ds-split-side" aria-labelledby="connector-grid-title">
          <span className="ds-eyebrow tool-eyebrow">Normalized bound range · {Math.max(0, result.grid.length - 1)} rows</span>
          <h2 id="connector-grid-title">Spreadsheet-ready grid</h2>
          <GridTable values={result.grid} label="Normalized connector data" />
          <p className="tool-note ds-muted">The same normalization rules back bound-range refreshes. Authentication, URL policy, and secret storage stay with the embedding host.</p>
          <div className="ds-panel">
            <span className="ds-eyebrow">Deterministic pipeline</span>
            <h2 id="connector-preprocess-title">Range preprocess</h2>
            <DsChip>{processed?.stages.length ?? 0} stages</DsChip>
            {processError && <p className="tool-error" role="alert">{processError}</p>}
            {processed && (
              <>
                <p className="tool-note ds-muted">{processed.stages.join(' → ')} · {processed.fingerprint}</p>
                <GridTable values={processed.grid} label="Preprocessed connector data" />
              </>
            )}
          </div>
        </section>
      </div>
    </section>
    </div>
  )
}

function GridTable({ values, label }: { values: unknown[][]; label: string }) {
  return (
    <div className="tool-table-wrap">
      <table className="tool-table ds-table">
        <caption className="visually-hidden">{label}</caption>
        <thead>
          <tr>
            {(values[0] ?? []).map((cell, index) => (
              <th key={index} className={typeof values[1]?.[index] === 'number' ? 'num' : undefined}>{String(cell ?? '')}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {values.slice(1).map((row, rowIndex) => (
            <tr key={rowIndex}>
              {row.map((cell, index) => (
                <td key={index} className={typeof cell === 'number' ? 'num' : undefined}>
                  {typeof cell === 'number' ? cell.toLocaleString() : String(cell ?? '')}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}
