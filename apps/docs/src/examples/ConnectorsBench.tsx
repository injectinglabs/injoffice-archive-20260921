import { useMemo, useState } from 'react'
import { csvToGrid, jsonToGrid } from '@injoffice/connectors'
import { LiveBench } from '../components/LiveBench'

const CSV = 'city,temp\nOslo,14\nLima,22\nKyoto,18'
const JSON_TEXT = '{ "results": [{ "city": "Oslo", "temp": 14 }, { "city": "Lima", "temp": 22 }] }'

export function ConnectorsBench() {
  const [mode, setMode] = useState<'csv' | 'json'>('csv')
  const [csv, setCsv] = useState(CSV)
  const [json, setJson] = useState(JSON_TEXT)
  const grid = useMemo(() => {
    if (mode === 'csv') return csvToGrid(csv)
    try {
      return jsonToGrid(JSON.parse(json) as unknown, 'results')
    } catch {
      return [['error', 'invalid JSON']]
    }
  }, [csv, json, mode])
  return (
    <LiveBench title="Live example" hint="@injoffice/connectors · csvToGrid / jsonToGrid">
      <div className="bench-controls">
        <label className="field">
          Source
          <select value={mode} onChange={(event) => setMode(event.target.value as 'csv' | 'json')}>
            <option value="csv">CSV</option>
            <option value="json">JSON path results</option>
          </select>
        </label>
        <label className="field" style={{ flex: 1 }}>
          Input
          <textarea rows={5} value={mode === 'csv' ? csv : json} onChange={(event) => mode === 'csv' ? setCsv(event.target.value) : setJson(event.target.value)} spellCheck={false} />
        </label>
      </div>
      <table className="grid-table">
        <tbody>
          {grid.map((row, rowIndex) => (
            <tr key={rowIndex}>
              {row.map((cell, columnIndex) => <td key={columnIndex}>{String(cell ?? '')}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
    </LiveBench>
  )
}
