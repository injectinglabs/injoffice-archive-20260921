import { Article } from '../components/Article'
import { Callout } from '../components/Callout'
import { CodeBlock } from '../components/CodeBlock'
import { Install } from '../components/Install'
import { ChartsBench } from '../examples/ChartsBench'
import { MutationsBench } from '../examples/MutationsBench'

export function IntroductionPage() {
  return (
    <Article id="introduction">
      <h2 id="what">What you get</h2>
      <p>InjOffice is seventeen TypeScript packages plus Go modules that extract, mutate, and verify Office artifacts. You can run entirely in the browser, call a local sidecar, or embed the packages in your own host. A hosted InjOffice service is not required.</p>
      <p>The original OOXML or PDF bytes stay authoritative. Editors, canvases, and Univer snapshots are optional views. If a change cannot be represented exactly, the API refuses it instead of approximating.</p>
      <h2 id="who">Who it is for</h2>
      <p>Application authors who need to open a real Excel, Word, PowerPoint, or PDF file, apply a bounded change, and return the same package with untouched parts byte-identical. Agents and web apps use the same JSON contracts.</p>
      <h2 id="not">What it is not</h2>
      <ul>
        <li>Microsoft Office parity</li>
        <li>A commercial Univer Pro license or Universer deployment</li>
        <li>An unrestricted XLSX rewriter</li>
      </ul>
      <Callout kind="note">The playground at <code>apps/playground</code> is the interactive proof of these packages. These docs teach you how to call the same APIs in your own app.</Callout>
    </Article>
  )
}

export function InstallationPage() {
  return (
    <Article id="installation">
      <h2 id="requirements">Requirements</h2>
      <p>Node.js 22 or newer for the TypeScript packages. Go, as declared in each module’s <code>go.mod</code>, for native extract/apply. React 18+ only if you use a UI integration.</p>
      <h2 id="packages">Install a package</h2>
      <p>Each package is independently consumable. Install only what you need:</p>
      <Install packages="@injoffice/charts echarts" />
      <Install packages="@injoffice/sheets" />
      <Install packages="@injoffice/univer-sheets @univerjs/core react react-dom rxjs" />
      <Install packages="@injoffice/xlsx-wasm" />
      <h2 id="workspace">This repository</h2>
      <CodeBlock language="bash" title="bash" code={`git clone https://github.com/injectinglabs/injoffice.git
cd injoffice
npm ci
npm run dev`} />
      <p><code>npm run dev</code> starts the workbench and these guides on port 3100. Open <code>#/guides</code> for documentation and <code>#/overview</code> for the interactive labs. Native extract/apply in the browser uses <code>@injoffice/xlsx-wasm</code>, <code>@injoffice/docx-wasm</code>, and <code>@injoffice/pptx-wasm</code>.</p>
      <h2 id="sidecar">Optional native sidecar</h2>
      <CodeBlock language="bash" title="bash" code={`cd go/injoffice-server
go run ./cmd/injoffice-server --addr 127.0.0.1:18765 --artifacts ./artifacts`} />
      <p>Vite proxies <code>/v1/xlsx</code>, <code>/v1/docx</code>, <code>/v1/pptx</code>, and <code>/v1/collab</code> to that process. Browser labs do not need it.</p>
    </Article>
  )
}

export function ConceptsPage() {
  return (
    <Article id="concepts">
      <h2 id="authority">File authority</h2>
      <p>Univer’s snapshot is a runtime document. InjOffice’s native JSON is a projection of an existing package. Saving always goes back through Go <code>Apply*</code> into the original archive. DOM, Canvas, screenshots, and Univer must not decide pass/fail.</p>
      <h2 id="pipeline">The native pipeline</h2>
      <ol>
        <li>Go <code>Extract*</code> reads package bytes.</li>
        <li>TypeScript decodes versioned native JSON.</li>
        <li>Optional UI (Univer, React, Konva) displays a view.</li>
        <li>Editors emit mutation JSON.</li>
        <li>Go <code>Apply*</code> writes those mutations into the original archive.</li>
      </ol>
      <h2 id="mutations">Mutations</h2>
      <p>Workbook changes use <code>injoffice.xlsx.mutations</code> v1: stable <code>sheet_id</code>, compare-and-swap <code>expected_revision</code>, and an idempotent <code>batch_id</code>. Unknown kinds are refused. Sheet add/delete/rename and row insert/delete are unsupported in v1.</p>
      <MutationsBench />
      <h2 id="univer-shell">Univer as a shell</h2>
      <p><code>@injoffice/univer-sheets</code> composes Apache-2.0 Univer presets. It does not import <code>@univerjs-pro/*</code>. Charts, pivots, shapes, collaboration, history, and native files are independent InjOffice packages that can hang off the grid.</p>
      <Callout kind="caution">A green ribbon item is not persistence proof. Round-tripping a native Office file is a separate claim, documented on the XLSX, DOCX, and PPTX pages.</Callout>
      <h2 id="fail-closed">Fail closed</h2>
      <p>Validation failures are JSON issues with a stable code, RFC 6901 path, and message. A refusal never returns a partial file.</p>
    </Article>
  )
}

export function QuickstartPage() {
  return (
    <Article id="quickstart">
      <h2 id="chart">1. Render a chart from a grid</h2>
      <Install packages="@injoffice/charts echarts react react-dom" />
      <CodeBlock
        language="ts"
        code={`import { buildEChartsOption, extractChartData, type ChartSpec } from '@injoffice/charts'

const values = [
  ['Month', 'Revenue'],
  ['Jan', 120],
  ['Feb', 165],
]
const spec: ChartSpec = {
  id: 'revenue',
  type: 'column',
  title: 'Revenue',
  range: { sheetId: 'sheet-1', startRow: 0, startColumn: 0, endRow: 2, endColumn: 1 },
}
const option = buildEChartsOption(spec, extractChartData(values, spec))`}
      />
      <ChartsBench />
      <h2 id="mutation">2. Validate a workbook mutation</h2>
      <Install packages="@injoffice/sheets" />
      <CodeBlock
        language="ts"
        code={`import { decodeWorkbookMutationBatch, encodeWorkbookMutationBatch } from '@injoffice/sheets'

const decoded = decodeWorkbookMutationBatch({
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
})
if (!decoded.ok) console.error(decoded.issues)
else await fetch('/v1/xlsx/mutations', { method: 'POST', body: encodeWorkbookMutationBatch(decoded.value) })`}
      />
      <h2 id="extract">3. Extract a real XLSX in the browser</h2>
      <Install packages="@injoffice/xlsx-wasm" />
      <CodeBlock
        language="ts"
        code={`import { createXlsxWasmClient } from '@injoffice/xlsx-wasm'

const original = new Uint8Array(await file.arrayBuffer())
const client = createXlsxWasmClient()
const workbook = await client.extract(original)
console.log(workbook.sheets.map((sheet) => sheet.name))
client.terminate()`}
      />
      <p>The Go sidecar on <code>127.0.0.1:18765</code> is an optional host fallback, not a requirement for this step. Next: <a href="#/guides/xlsx">Native XLSX</a> or <a href="#/guides/sheets-editor">the Univer editor shell</a>.</p>
    </Article>
  )
}
