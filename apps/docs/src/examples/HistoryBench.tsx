import { useMemo, useState } from 'react'
import { diffGrids } from '@injoffice/history'
import { LiveBench } from '../components/LiveBench'

const BEFORE = 'Name,Plan,Actual\nNorth,120,118\nSouth,95,101\nWest,140,139'
const AFTER = 'Name,Plan,Actual\nNorth,120,126\nSouth,105,101\nWest,140,151\nEast,80,84'

function parse(value: string): string[][] {
  return value.split('\n').map((row) => row.split(',').map((cell) => cell.trim()))
}

export function HistoryBench() {
  const [before, setBefore] = useState(BEFORE)
  const [after, setAfter] = useState(AFTER)
  const diff = useMemo(() => diffGrids(parse(before), parse(after)), [after, before])
  return (
    <LiveBench title="Live example" hint="@injoffice/history · diffGrids">
      <div className="bench-controls">
        <label className="field" style={{ flex: 1 }}>
          Before
          <textarea rows={5} value={before} onChange={(event) => setBefore(event.target.value)} />
        </label>
        <label className="field" style={{ flex: 1 }}>
          After
          <textarea rows={5} value={after} onChange={(event) => setAfter(event.target.value)} />
        </label>
      </div>
      <p>{diff.changeCount} cell change{diff.changeCount === 1 ? '' : 's'}</p>
      <table className="grid-table">
        <thead><tr><th>Cell</th><th>From</th><th>To</th></tr></thead>
        <tbody>
          {diff.changes.map((change) => (
            <tr key={change.address}><td>{change.address}</td><td>{change.from}</td><td>{change.to}</td></tr>
          ))}
        </tbody>
      </table>
    </LiveBench>
  )
}
