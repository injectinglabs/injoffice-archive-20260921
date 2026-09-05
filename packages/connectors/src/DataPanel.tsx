import { useEffect, useReducer, useState } from 'react'
import type { ConnectorManager } from './manager'
import type { ConnectorSpec } from './types'

// DataPanel — the connectors UI: list bound ranges with refresh status, add
// a new HTTP source at the current selection, refresh on demand. Same ioc-*
// styling hooks as the chart/pivot panels.

function timeAgo(ts: number): string {
  const s = Math.max(1, Math.round((Date.now() - ts) / 1000))
  if (s < 60) return `${s}s ago`
  const m = Math.round(s / 60)
  if (m < 60) return `${m}m ago`
  return `${Math.round(m / 60)}h ago`
}

export function DataPanel({ manager }: { manager: ConnectorManager }) {
  const [, force] = useReducer((x: number) => x + 1, 0)
  const [adding, setAdding] = useState(false)
  const [url, setUrl] = useState('')
  const [format, setFormat] = useState<'json' | 'csv'>('json')
  const [path, setPath] = useState('')
  const [name, setName] = useState('')

  useEffect(() => manager.onChange(force), [manager])

  const connectors = manager.list()

  const submit = async () => {
    if (!url.trim()) return
    await manager.createAtSelection({
      name: name.trim() || new URL(url, 'https://x').hostname || 'Connection',
      source: { kind: 'http', url: url.trim(), format, path: path.trim() || undefined },
      refresh: 'manual',
    })
    setAdding(false)
    setUrl('')
    setPath('')
    setName('')
  }

  return (
    <div className="ioc-panel">
      <div className="ioc-series-head">Data connections</div>
      {connectors.length === 0 && !adding && (
        <div className="ioc-panel--empty">
          No connections yet — select a target cell, then connect a source. The
          fetched table lands there and refreshes on demand.
        </div>
      )}
      {connectors.map((c: ConnectorSpec) => {
        const st = manager.status(c.id)
        return (
          <div key={c.id} className="ioc-conn">
            <div className="ioc-conn-row">
              <span className="ioc-conn-name" title={c.source.url}>{c.name}</span>
              <button type="button" disabled={st?.refreshing} onClick={() => void manager.refresh(c.id)}>
                {st?.refreshing ? '…' : 'Refresh'}
              </button>
              <button type="button" className="ioc-remove" onClick={() => manager.remove(c.id)}>
                ×
              </button>
            </div>
            <div className="ioc-conn-status">
              {st?.lastError
                ? <span className="ioc-conn-err">{st.lastError}</span>
                : st?.lastRefreshTs
                  ? `${st.rowCount ?? 0}×${st.columnCount ?? 0} · ${timeAgo(st.lastRefreshTs)}`
                  : 'not fetched yet'}
            </div>
          </div>
        )
      })}
      {adding ? (
        <div className="ioc-conn-form">
          <input placeholder="Name (optional)" value={name} onChange={(e) => setName(e.target.value)} />
          <input placeholder="https://… (JSON or CSV)" value={url} onChange={(e) => setUrl(e.target.value)} />
          <div className="ioc-row">
            <select aria-label="Format" value={format} onChange={(e) => setFormat(e.target.value as 'json' | 'csv')}>
              <option value="json">JSON</option>
              <option value="csv">CSV</option>
            </select>
            {format === 'json' && (
              <input placeholder="path (e.g. data.items)" value={path} onChange={(e) => setPath(e.target.value)} />
            )}
          </div>
          <div className="ioc-row">
            <button type="button" onClick={() => void submit()} disabled={!url.trim()}>
              Connect
            </button>
            <button type="button" onClick={() => setAdding(false)}>Cancel</button>
          </div>
        </div>
      ) : (
        <button type="button" className="ioc-add" onClick={() => setAdding(true)}>
          + connection
        </button>
      )}
    </div>
  )
}
