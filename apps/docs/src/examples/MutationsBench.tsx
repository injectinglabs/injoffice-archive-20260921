import { useMemo, useState } from 'react'
import { decodeWorkbookMutationBatch } from '@injoffice/sheets'
import { LiveBench } from '../components/LiveBench'

const SAMPLE = `{
  "protocol": "injoffice.xlsx.mutations",
  "version": 1,
  "batch_id": "save-0001",
  "expected_revision": "rev:demo",
  "operations": [
    {
      "operation_id": "edit-0001",
      "kind": "cell.set_value",
      "sheet_id": "1",
      "cell": { "row": 1, "column": 2 },
      "value": 42
    }
  ]
}`

export function MutationsBench() {
  const [text, setText] = useState(SAMPLE)
  const result = useMemo(() => {
    try {
      return decodeWorkbookMutationBatch(JSON.parse(text) as unknown)
    } catch (error) {
      return { ok: false as const, issues: [{ code: 'INVALID_VALUE' as const, message: error instanceof Error ? error.message : 'Invalid JSON', path: '/' }] }
    }
  }, [text])
  return (
    <LiveBench title="Live example" hint="@injoffice/sheets · decodeWorkbookMutationBatch">
      <div className="bench-controls">
        <label className="field" style={{ flex: 1 }}>
          Mutation batch
          <textarea rows={12} value={text} onChange={(event) => setText(event.target.value)} spellCheck={false} />
        </label>
      </div>
      {result.ok
        ? <p><span className="badge badge--pass">valid</span> {result.value.operations.length} operation{result.value.operations.length === 1 ? '' : 's'}</p>
        : (
          <div>
            <p><span className="badge badge--patch">refused</span></p>
            <ul>{result.issues.map((issue, index) => <li key={index}><code>{issue.code}</code> {issue.message}</li>)}</ul>
          </div>
        )}
    </LiveBench>
  )
}
