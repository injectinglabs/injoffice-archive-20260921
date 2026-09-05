import { useMemo, useState } from 'react'
import { AGG_KINDS, bakePivot, fieldValues, type AggKind } from '@injoffice/pivots'
import { playgroundPivotSpec, playgroundPivotWire } from '../pivotWire'

const SOURCE: (string | number)[][] = [
  ['Region', 'Owner', 'Quarter', 'Revenue', 'Units'],
  ['West', 'Maya', 'Q1', 42000, 18],
  ['East', 'Theo', 'Q1', 37000, 15],
  ['West', 'Maya', 'Q2', 51000, 21],
  ['East', 'Inez', 'Q2', 46000, 19],
  ['North', 'Theo', 'Q1', 29000, 12],
  ['North', 'Inez', 'Q2', 34000, 14],
]

export default function PivotsPage() {
  const [rowField, setRowField] = useState('Region')
  const [valueField, setValueField] = useState('Revenue')
  const [aggregation, setAggregation] = useState<AggKind>('sum')
  const [splitQuarter, setSplitQuarter] = useState(true)
  const [region, setRegion] = useState('All')

  const regions = useMemo(() => fieldValues(SOURCE, 'Region'), [])
  const pivot = useMemo(() => bakePivot(SOURCE, {
    rows: [rowField],
    columns: splitQuarter ? ['Quarter'] : [],
    values: [{ field: valueField, agg: aggregation }],
    filters: region === 'All' ? undefined : { Region: [region] },
    grandTotals: true,
  }), [aggregation, region, rowField, splitQuarter, valueField])
  const native = useMemo(() => {
    const spec = playgroundPivotSpec({
      rows: [rowField],
      columns: splitQuarter ? ['Quarter'] : [],
      valueField,
      aggregation,
      filters: region === 'All' ? undefined : { Region: [region] },
      rowCount: SOURCE.length,
      columnCount: SOURCE[0]!.length,
    })
    return playgroundPivotWire(spec, SOURCE[0]!, {
      Region: fieldValues(SOURCE, 'Region').map(String),
      Quarter: fieldValues(SOURCE, 'Quarter').map(String),
    })
  }, [aggregation, region, rowField, splitQuarter, valueField])

  return (
    <section className="tool-page" data-demo-surface="pivots" aria-label="Pivot table workbench">
      <div className="tool-page__controls" role="toolbar" aria-label="Pivot controls">
        <label className="tool-field">Group rows<select value={rowField} onChange={(event) => setRowField(event.target.value)}><option>Region</option><option>Owner</option><option>Quarter</option></select></label>
        <label className="tool-field">Value<select value={valueField} onChange={(event) => setValueField(event.target.value)}><option>Revenue</option><option>Units</option></select></label>
        <label className="tool-field">Aggregation<select value={aggregation} onChange={(event) => setAggregation(event.target.value as AggKind)}>{AGG_KINDS.map((kind) => <option key={kind}>{kind}</option>)}</select></label>
        <label className="tool-field">Region<select value={region} onChange={(event) => setRegion(event.target.value)}><option>All</option>{regions.map((value) => <option key={value}>{value}</option>)}</select></label>
        <label className="tool-check"><input type="checkbox" checked={splitQuarter} onChange={(event) => setSplitQuarter(event.target.checked)} />Split by quarter</label>
      </div>

      <div className="tool-page__grid">
        <section className="tool-card" aria-labelledby="pivot-source-title">
          <div className="tool-card__heading"><div><span className="tool-eyebrow">Input</span><h2 id="pivot-source-title">Sales records</h2></div><span className="tool-chip">{SOURCE.length - 1} rows</span></div>
          <DataTable values={SOURCE} label="Sales source data" />
        </section>
        <section className="tool-card tool-card--hero" aria-labelledby="pivot-output-title">
          <div className="tool-card__heading"><div><span className="tool-eyebrow">Live result</span><h2 id="pivot-output-title">Generated pivot</h2></div><span className="tool-chip">{pivot.rowCount} × {pivot.columnCount}</span></div>
          <DataTable values={pivot.grid} label="Generated pivot table" />
          <p className="tool-note">Filters, totals, row groups, column groups, and aggregation are represented by a plain-JSON PivotSpec.</p>
          <p className="tool-note">{native.representability.representable ? 'Native conversion is representable.' : native.representability.issues.map((issue) => issue.message).join(' ')}</p>
          <pre className="tool-note">{JSON.stringify(native.wire.pivots[0] ?? { skipped: native.wire.skipped }, null, 2)}</pre>
        </section>
      </div>
    </section>
  )
}

function DataTable({ values, label }: { values: unknown[][]; label: string }) {
  return (
    <div className="tool-table-wrap">
      <table className="tool-table">
        <caption className="visually-hidden">{label}</caption>
        <thead><tr>{(values[0] ?? []).map((cell, index) => <th key={index}>{String(cell ?? '')}</th>)}</tr></thead>
        <tbody>{values.slice(1).map((row, rowIndex) => <tr key={rowIndex}>{row.map((cell, index) => <td key={index}>{typeof cell === 'number' ? cell.toLocaleString() : String(cell ?? '')}</td>)}</tr>)}</tbody>
      </table>
    </div>
  )
}
