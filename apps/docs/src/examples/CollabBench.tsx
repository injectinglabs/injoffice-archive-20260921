import { useMemo, useState } from 'react'
import { selectionA1, transformSelection, type SheetSelection, type StructEdit } from '@injoffice/collab'
import { LiveBench } from '../components/LiveBench'

export function CollabBench() {
  const [row, setRow] = useState(2)
  const [insertAt, setInsertAt] = useState(0)
  const [count, setCount] = useState(1)
  const selection = useMemo<SheetSelection>(() => ({
    sheet: 's1',
    ranges: [[row, 0, row, 0]],
    active: [row, 0],
  }), [row])
  const edit = useMemo<StructEdit>(() => ({
    kind: 'insert',
    axis: 'row',
    subUnitId: 's1',
    start: insertAt,
    count,
  }), [count, insertAt])
  const next = useMemo(() => transformSelection(selection, edit, 's1'), [edit, selection])
  return (
    <LiveBench title="Live example" hint="@injoffice/collab · transformSelection">
      <div className="bench-controls">
        <label className="field">
          Selection row
          <input type="number" min={0} value={row} onChange={(event) => setRow(Number(event.target.value))} />
        </label>
        <label className="field">
          Insert rows at
          <input type="number" min={0} value={insertAt} onChange={(event) => setInsertAt(Number(event.target.value))} />
        </label>
        <label className="field">
          Count
          <input type="number" min={1} value={count} onChange={(event) => setCount(Math.max(1, Number(event.target.value)))} />
        </label>
      </div>
      <p>Before <code>{selectionA1(selection) || '(none)'}</code> → after <code>{selectionA1(next) || '(consumed)'}</code></p>
      <p>The package still needs a host transport for rooms. This bench is the pure transform used when a peer inserts rows ahead of a local selection.</p>
    </LiveBench>
  )
}
