import { useMemo, useState } from 'react'
import { AGG_KINDS, bakePivot, type AggKind } from '@injoffice/pivots'
import { LiveBench } from '../components/LiveBench'

const SOURCE: (string | number)[][] = [
  ['Region', 'Owner', 'Quarter', 'Revenue'],
  ['West', 'Maya', 'Q1', 42000],
  ['East', 'Theo', 'Q1', 37000],
  ['West', 'Maya', 'Q2', 51000],
  ['East', 'Inez', 'Q2', 46000],
  ['North', 'Theo', 'Q1', 29000],
  ['North', 'Inez', 'Q2', 34000],
]

export function PivotsBench() {
  const [rowField, setRowField] = useState('Region')
  const [aggregation, setAggregation] = useState<AggKind>('sum')
  const [splitQuarter, setSplitQuarter] = useState(true)
  const pivot = useMemo(() => bakePivot(SOURCE, {
    rows: [rowField],
    columns: splitQuarter ? ['Quarter'] : [],
    values: [{ field: 'Revenue', agg: aggregation }],
    grandTotals: true,
  }), [aggregation, rowField, splitQuarter])
  return (
    <LiveBench title="Live example" hint="@injoffice/pivots · bakePivot">
      <div className="bench-controls">
        <label className="field">Rows<select value={rowField} onChange={(event) => setRowField(event.target.value)}><option>Region</option><option>Owner</option></select></label>
        <label className="field">Aggregation<select value={aggregation} onChange={(event) => setAggregation(event.target.value as AggKind)}>{AGG_KINDS.map((kind) => <option key={kind}>{kind}</option>)}</select></label>
        <label className="field"><span>Columns</span><label><input type="checkbox" checked={splitQuarter} onChange={(event) => setSplitQuarter(event.target.checked)} /> Split by quarter</label></label>
      </div>
      <table className="grid-table">
        <tbody>
          {pivot.grid.map((row, rowIndex) => (
            <tr key={rowIndex}>
              {row.map((cell, columnIndex) => rowIndex === 0
                ? <th key={columnIndex}>{cell ?? ''}</th>
                : <td key={columnIndex}>{cell ?? ''}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
    </LiveBench>
  )
}
