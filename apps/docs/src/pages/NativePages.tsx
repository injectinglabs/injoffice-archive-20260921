import { lazy, Suspense } from 'react'
import { Article } from '../components/Article'
import { Callout } from '../components/Callout'
import { CodeBlock } from '../components/CodeBlock'
import { Install } from '../components/Install'
import { MutationsBench } from '../examples/MutationsBench'

const NativeExtractBench = lazy(() => import('../examples/NativeExtractBench').then((mod) => ({ default: mod.NativeExtractBench })))

export function XlsxPage() {
  return (
    <Article id="xlsx">
      <h2 id="install">Installation</h2>
      <Install packages="@injoffice/sheets" />
      <h2 id="decode">Decode a mutation batch</h2>
      <p>Version 1 normalizes editor changes into a small JSON vocabulary. It does not apply bytes. Coordinates are zero-based. <code>sheet_id</code> is the stable identity from the workbook, never a tab name.</p>
      <CodeBlock
        language="ts"
        code={`import { decodeWorkbookMutationBatch, encodeWorkbookMutationBatch } from '@injoffice/sheets'

const batch = {
  protocol: 'injoffice.xlsx.mutations',
  version: 1,
  batch_id: 'save-0001',
  expected_revision: 'opaque-revision-from-gateway',
  operations: [{
    operation_id: 'edit-0001',
    kind: 'cell.set_value',
    sheet_id: 'stable-sheet-id',
    cell: { row: 1, column: 2 },
    value: 42,
  }],
} as const

const decoded = decodeWorkbookMutationBatch(batch)
if (!decoded.ok) console.error(decoded.issues)
else {
  const body = encodeWorkbookMutationBatch(decoded.value)
  await fetch('/v1/xlsx/mutations', { method: 'POST', body })
}`}
      />
      <MutationsBench />
      <h2 id="kinds">Supported operations</h2>
      <table>
        <thead><tr><th>Kind</th><th>Effect</th></tr></thead>
        <tbody>
          <tr><td><code>cell.set_value</code></td><td>Literal write; removes any formula</td></tr>
          <tr><td><code>cell.set_formula</code></td><td>A1 formula including <code>=</code>; removes the literal</td></tr>
          <tr><td><code>cell.clear_value</code> / <code>cell.clear_formula</code></td><td>Leave the cell blank</td></tr>
          <tr><td><code>style.patch</code></td><td>Omitted keys unchanged; <code>null</code> clears</td></tr>
          <tr><td><code>range.merge</code> / <code>range.unmerge</code></td><td>The only structural kinds in v1</td></tr>
        </tbody>
      </table>
      <Callout kind="caution">Sheet add/delete/rename/reorder and row/column insert/delete/move are exported as <code>UNSUPPORTED_STRUCTURAL_MUTATION_KINDS</code>. Validation returns <code>UNSUPPORTED_OPERATION</code> instead of dropping them.</Callout>
      <h2 id="extract">Extract and apply in the browser</h2>
      <Install packages="@injoffice/xlsx-wasm" />
      <p>The published WASM client runs the same Go engine as the sidecar, inside a Worker. No HTTP service is required.</p>
      <CodeBlock
        language="ts"
        code={`import { adaptWorkbookMutationBatchV1, createXlsxWasmClient } from '@injoffice/xlsx-wasm'

const original = new Uint8Array(await file.arrayBuffer())
const client = createXlsxWasmClient()
const workbook = await client.extract(original)
const saved = await client.apply(original, workbook, adaptWorkbookMutationBatchV1(workbook, batch))
client.terminate()`}
      />
      <Suspense fallback={<p>Loading the browser XLSX engine…</p>}>
        <NativeExtractBench />
      </Suspense>
      <p><code>ExtractNativeWorkbookV2</code> is the read projection (geometry, decorations, unsupported content preserved exactly). v1 remains the mutation/save contract. Formula text is never evaluated.</p>
      <h2 id="sidecar">Optional sidecar</h2>
      <p>Use the sidecar when you already run Go on a host. Browser labs do not need it.</p>
      <CodeBlock
        language="bash"
        code={`cd go/xlsxpatch
go run ./cmd/xlsxnative serve

curl -sS -H 'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' \\
  --data-binary @input.xlsx http://127.0.0.1:18765/v1/xlsx/extract`}
      />
      <h2 id="go">Go extract</h2>
      <CodeBlock
        language="go"
        code={`native, err := xlsxpatch.ExtractNativeWorkbookV1(input)
if err != nil { panic(err) }
next, err := xlsxpatch.ExtractNativeWorkbookV1WithOptions(updatedInput,
  xlsxpatch.NativeWorkbookExtractionOptions{Previous: native})`}
      />
    </Article>
  )
}

export function DocxPage() {
  return (
    <Article id="docx">
      <h2 id="install">Installation</h2>
      <Install packages="@injoffice/docs" />
      <h2 id="decode">Decode native JSON</h2>
      <CodeBlock
        language="ts"
        code={`import { decodeNativeDocxJson, encodeNativeDocxDocument } from '@injoffice/docs'

const decoded = decodeNativeDocxJson(await response.text())
if (!decoded.ok) console.error(decoded.issues)
else {
  const canonical = encodeNativeDocxDocument(decoded.value)
}`}
      />
      <p>The published schema is <code>@injoffice/docs/native-docx-v1.schema.json</code>. Unknown fields, dangling identities, and over-limit inputs are rejected.</p>
      <h2 id="run">Guarded run replacement</h2>
      <p>v1 accepts only a closed text replace wholly inside one native text run. Tables, drawings, notes, fields, and tracked markup return issues with no envelope.</p>
      <Install packages="@injoffice/docx-wasm" />
      <CodeBlock
        language="ts"
        code={`import { createDocxWasmClient } from '@injoffice/docx-wasm'

const original = new Uint8Array(await file.arrayBuffer())
const client = createDocxWasmClient()
const document = await client.extract(original)
const saved = await client.apply(original, document, envelope)
client.terminate()`}
      />
      <p>The sidecar is optional fallback, not the default:</p>
      <CodeBlock
        language="bash"
        code={`curl -sS --data-binary @input.docx http://127.0.0.1:18765/v1/docx/extract
# POST /v1/docx/mutations with original + payload + expected_revision`}
      />
      <Callout kind="caution">The playground docs surface is not Word pagination or page-paint. Page-paint needs an exact font/layout bundle that public extract does not yet return.</Callout>
    </Article>
  )
}

export function PptxPage() {
  return (
    <Article id="pptx">
      <h2 id="native">Native contract</h2>
      <Install packages="@injoffice/pptx-native" />
      <p>Versioned native PPTX JSON is shared by parsers, patchers, and renderers. Browser extract/apply is <code>@injoffice/pptx-wasm</code>; <code>go/pptxpatch</code> is the same engine on a host.</p>
      <Install packages="@injoffice/pptx-wasm" />
      <CodeBlock
        language="ts"
        code={`import { createPptxWasmClient } from '@injoffice/pptx-wasm'

const original = new Uint8Array(await file.arrayBuffer())
const client = createPptxWasmClient()
const deck = await client.extract(original)
const saved = await client.apply(original, deck, mutation)
client.terminate()`}
      />
      <h2 id="authored">Compile an authored deck</h2>
      <Install packages="@injoffice/pptx-authored" />
      <CodeBlock
        language="ts"
        code={`import { compileDeckSpecToNativeV1 } from '@injoffice/pptx-authored'

const result = compileDeckSpecToNativeV1(spec)
if (!result.ok) throw new Error(result.issues.map(({ message }) => message).join('; '))
const nativeDeck = result.deck`}
      />
      <p>Compilation is atomic. Unknown fields and unmodeled wire semantics return a refusal with no partial deck. HTML/CSS measurement is not an authoring authority.</p>
      <h2 id="http">Optional HTTP sidecar</h2>
      <CodeBlock
        language="bash"
        code={`curl -sS --data-binary @deck.pptx http://127.0.0.1:18765/v1/pptx/extract`}
      />
    </Article>
  )
}
