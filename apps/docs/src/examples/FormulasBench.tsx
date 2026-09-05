import { useMemo, useState } from 'react'
import { TARGET_FUNCTIONS } from '@injoffice/formulas'
import { LiveBench } from '../components/LiveBench'

export function FormulasBench() {
  const [query, setQuery] = useState('XLOOKUP')
  const matches = useMemo(() => {
    const needle = query.trim().toLowerCase()
    return TARGET_FUNCTIONS.filter((item) => item.name.toLowerCase().includes(needle) || item.formula.toLowerCase().includes(needle)).slice(0, 8)
  }, [query])
  return (
    <LiveBench title="Live example" hint="@injoffice/formulas · TARGET_FUNCTIONS">
      <div className="bench-controls">
        <label className="field" style={{ flex: 1 }}>
          Search audited functions
          <input value={query} onChange={(event) => setQuery(event.target.value)} />
        </label>
        <span className="badge">{TARGET_FUNCTIONS.length} in the matrix</span>
      </div>
      <table className="grid-table">
        <thead><tr><th>Name</th><th>Valid fixture invocation</th></tr></thead>
        <tbody>
          {matches.map((item) => (
            <tr key={item.name}><td>{item.name}</td><td>{item.formula}</td></tr>
          ))}
        </tbody>
      </table>
    </LiveBench>
  )
}
