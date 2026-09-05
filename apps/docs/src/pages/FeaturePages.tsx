import { lazy, Suspense } from 'react'
import { Article } from '../components/Article'
import { Callout } from '../components/Callout'
import { CodeBlock } from '../components/CodeBlock'
import { Install } from '../components/Install'
import { ChartsBench } from '../examples/ChartsBench'
import { ConnectorsBench } from '../examples/ConnectorsBench'
import { FormulasBench } from '../examples/FormulasBench'
import { HistoryBench } from '../examples/HistoryBench'
import { PivotsBench } from '../examples/PivotsBench'
import { ShapesBench } from '../examples/ShapesBench'

const EditorBench = lazy(() => import('../examples/EditorBench').then((mod) => ({ default: mod.EditorBench })))

export function SheetsEditorPage() {
  return (
    <Article id="sheets-editor">
      <h2 id="install">Installation</h2>
      <Install packages="@injoffice/univer-sheets @univerjs/core react react-dom rxjs" />
      <h2 id="preset">Preset composition</h2>
      <p>This is InjOffice’s equivalent of Univer preset mode: a curated Apache-2.0 plugin set, without <code>@univerjs/presets</code> (that umbrella pulls Pro packages).</p>
      <CodeBlock
        language="ts"
        code={`import { LocaleType } from '@univerjs/core'
import {
  createOssUniver,
  createSheetsPresetBundle,
  registerInjOfficeInsertMenu,
} from '@injoffice/univer-sheets/browser'
import '@injoffice/univer-sheets/styles.css'

const bundle = createSheetsPresetBundle({
  container: 'app',
  features: {
    threadComments: false,
    notes: false,
  },
})
const { univer, univerAPI } = createOssUniver({
  locale: LocaleType.EN_US,
  locales: { [LocaleType.EN_US]: bundle.locale },
  presets: bundle.presets,
})
univerAPI.createWorkbook({})
registerInjOfficeInsertMenu(univer, {
  chart: () => chartManager.createFromSelection('column'),
  pivotTable: () => pivotManager.createFromSelection(),
  shape: () => shapeManager.create('rect'),
})`}
      />
      <Suspense fallback={<p>Loading Univer OSS editor…</p>}>
        <EditorBench />
      </Suspense>
      <h2 id="features">Public Univer surface</h2>
      <table>
        <thead><tr><th>Capability</th><th>Config key</th></tr></thead>
        <tbody>
          <tr><td>Core editing, formulas, number formats</td><td>mandatory</td></tr>
          <tr><td>Conditional formatting</td><td><code>conditionalFormatting</code></td></tr>
          <tr><td>Data validation</td><td><code>dataValidation</code></td></tr>
          <tr><td>Drawing and images</td><td><code>drawing</code></td></tr>
          <tr><td>Filters, sort, find/replace, hyperlinks, notes, tables, comments</td><td>matching keys, default on</td></tr>
        </tbody>
      </table>
      <Callout kind="caution">Setting a key to <code>false</code> omits the plugin and its ribbon items. Univer remains a snapshot UI. Persist through native mutations, not <code>workbook.save()</code>.</Callout>
    </Article>
  )
}

export function ChartsPage() {
  return (
    <Article id="charts">
      <h2 id="install">Installation</h2>
      <Install packages="@injoffice/charts echarts react react-dom" />
      <h2 id="usage">Usage</h2>
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
      <h2 id="types">Chart types</h2>
      <p>column, bar, line, area, pie, doughnut, scatter, waterfall, heatmap, radar, funnel, treemap, sankey, boxplot.</p>
      <h2 id="editor">Optional Univer integration</h2>
      <p><code>ChartManager</code>, <code>ChartFloat</code>, and <code>ChartPanel</code> mount on the OSS grid. Persistence of chart parts in XLSX is a host/Go concern. This is not <code>@univerjs-pro/sheets-chart</code>.</p>
      <Callout kind="note">Unsupported Excel plot types, effects, and surgical update/delete are documented as partial. The spec is the contract; ECharts is a renderer.</Callout>
    </Article>
  )
}

export function PivotsPage() {
  return (
    <Article id="pivots">
      <h2 id="install">Installation</h2>
      <Install packages="@injoffice/pivots react react-dom" />
      <h2 id="usage">Usage</h2>
      <CodeBlock
        language="ts"
        code={`import { bakePivot } from '@injoffice/pivots'

const result = bakePivot(
  [
    ['Region', 'Sales'],
    ['West', 20],
    ['East', 35],
    ['West', 15],
  ],
  { rows: ['Region'], columns: [], values: [{ field: 'Sales', agg: 'sum' }] },
)
console.log(result.grid)`}
      />
      <PivotsBench />
      <p>Aggregations: <code>sum</code>, <code>count</code>, <code>avg</code>, <code>min</code>, <code>max</code>. <code>PivotManager</code> is the optional Univer panel. <code>pivotsFromFile</code> converts native XLSX pivot JSON and refuses features it cannot represent.</p>
    </Article>
  )
}

export function ShapesPage() {
  return (
    <Article id="shapes">
      <h2 id="install">Installation</h2>
      <Install packages="@injoffice/shapes react react-dom" />
      <h2 id="usage">Usage</h2>
      <CodeBlock
        language="ts"
        code={`import { roundRectPath, shapeLabel } from '@injoffice/shapes'

const path = roundRectPath(240, 100, 0.16)
console.log(shapeLabel('roundRect'), path)`}
      />
      <ShapesBench />
      <p>Kind names are ECMA-376 preset geometry tokens (<code>a:prstGeom</code>). Unknown XLSX drawings stay under the host’s fail-closed preservation policy.</p>
    </Article>
  )
}

export function ConnectorsPage() {
  return (
    <Article id="connectors">
      <h2 id="install">Installation</h2>
      <Install packages="@injoffice/connectors" />
      <h2 id="usage">Usage</h2>
      <CodeBlock
        language="ts"
        code={`import { csvToGrid, jsonToGrid } from '@injoffice/connectors'

const rows = jsonToGrid({ results: [{ city: 'Oslo', temp: 14 }, { city: 'Lima', temp: 22 }] }, 'results')
const csvRows = csvToGrid('city,temp\\nOslo,14\\nLima,22')`}
      />
      <ConnectorsBench />
      <Callout kind="ok">Applications supply a <code>SourceFetcher</code> for auth, URL policy, and secrets. The package never stores credentials or performs arbitrary network requests on its own.</Callout>
    </Article>
  )
}

export function FormulasPage() {
  return (
    <Article id="formulas">
      <h2 id="install">Installation</h2>
      <Install packages="@injoffice/formulas" />
      <h2 id="usage">Usage</h2>
      <CodeBlock
        language="ts"
        code={`import { TARGET_FUNCTIONS } from '@injoffice/formulas'

for (const fn of TARGET_FUNCTIONS) {
  console.log(fn.name, fn.formula)
}`}
      />
      <FormulasBench />
      <p>This package is not an evaluator. CI evaluates each fixture invocation against the pinned Univer OSS engine and fails on <code>#NAME?</code> or computation regressions.</p>
    </Article>
  )
}

export function CollabPage() {
  return (
    <Article id="collab">
      <h2 id="install">Installation</h2>
      <Install packages="@injoffice/collab" />
      <h2 id="transport">Host transport</h2>
      <CodeBlock
        language="ts"
        code={`import { COLLAB_EVENTS, type CollabTransport } from '@injoffice/collab'

const transport: CollabTransport = {
  join: (path, name) => rpc('collab.join', { path, name }),
  leave: (path) => rpc('collab.leave', { path }),
  presence: (path, selection) => rpc('collab.presence', { path, selection }),
  onEvent: (handler) => socketEvents((frame) => {
    if (COLLAB_EVENTS.has(frame.event)) handler(frame)
  }),
  onReconnect: (handler) => reconnectEvents(handler),
  opSubmit: (path, ops, baseSeq) => rpc('collab.op.submit', { path, ops, base_seq: baseSeq }),
  opSince: (path, sinceSeq) => rpc('collab.op.since', { path, since_seq: sinceSeq }),
}`}
      />
      <p>The package includes no server. Optional in-repo hub:</p>
      <CodeBlock
        language="bash"
        code={`cd go/injoffice-server
go run ./cmd/injoffice-server --addr 127.0.0.1:18765 --artifacts ./artifacts

# POST /v1/collab/session
# GET  /v1/collab/events?session_id=...
# POST /v1/collab/join | presence | op/submit | op/since`}
      />
      <Callout kind="caution">This is not Univer Pro OT. The playground’s default demo is two independent Univer cores plus InjOffice presence/ops. Move/reorder transforms and host undo-stack rebasing remain gaps.</Callout>
    </Article>
  )
}

export function HistoryPage() {
  return (
    <Article id="history">
      <h2 id="install">Installation</h2>
      <Install packages="@injoffice/history" />
      <h2 id="usage">Usage</h2>
      <CodeBlock
        language="ts"
        code={`import { diffGrids, diffText } from '@injoffice/history'

const sheetDiff = diffGrids([['Name', 'Score'], ['Ada', 8]], [['Name', 'Score'], ['Ada', 10]])
const textDiff = diffText('draft one', 'draft two')`}
      />
      <HistoryBench />
      <p>Grid diffs are positional, not LCS row-alignment. Version storage, authors, and retention stay with the host. This is not Univer Pro edit-history restore.</p>
    </Article>
  )
}

export function FontMetricsPage() {
  return (
    <Article id="font-metrics">
      <h2 id="install">Installation</h2>
      <Install packages="@injoffice/font-metrics" />
      <h2 id="node">Node font discovery</h2>
      <CodeBlock
        language="ts"
        code={`import { findSystemFont, isFamilyInstalled } from '@injoffice/font-metrics'

if (isFamilyInstalled('Arial')) {
  const bytes = findSystemFont('ArialMT', 'Arial')
  if (bytes) console.log(\`found \${bytes.length} font bytes\`)
}`}
      />
      <h2 id="layout">Layout contract</h2>
      <CodeBlock
        language="ts"
        code={`import {
  NATIVE_TEXT_LAYOUT_VERSION,
  classifyNativeOfficeLineBreak,
  validateTextRunInput,
} from '@injoffice/font-metrics/layout'`}
      />
      <p><code>@injoffice/font-metrics/layout</code> is pure TypeScript: no DOM, canvas, or Node. Callers remain responsible for font licensing.</p>
    </Article>
  )
}

export function PdfGuidePage() {
  return (
    <Article id="pdf">
      <h2 id="install">Installation</h2>
      <Install packages="@injoffice/pdf" />
      <h2 id="usage">Usage</h2>
      <CodeBlock
        language="ts"
        code={`import { applyPageOps, readInfo } from '@injoffice/pdf/browser'

const source = new Uint8Array(await file.arrayBuffer())
const info = await readInfo(source)
const rotated = await applyPageOps(source, [
  { type: 'rotate', pages: [info.pageCount], degrees: 90 },
])`}
      />
      <CodeBlock
        language="ts"
        title="Vite worker"
        code={`import workerURL from 'pdfjs-dist/legacy/build/pdf.worker.min.mjs?url'
import { configurePdfWorker, PdfViewerDocument } from '@injoffice/pdf/browser'

configurePdfWorker(workerURL)
const viewer = await PdfViewerDocument.load(source)`}
      />
      <p>Univer’s OSS and Pro demos do not include a PDF engine. Treat transformed files as untrusted input.</p>
    </Article>
  )
}

export function SlidesGuidePage() {
  return (
    <Article id="slides">
      <h2 id="install">Installation</h2>
      <Install packages="@injoffice/slides @injoffice/pptx-authored @injoffice/pptx-render" />
      <h2 id="usage">Usage</h2>
      <CodeBlock
        language="ts"
        code={`import { compileDeckSpecToNativeV1 } from '@injoffice/pptx-authored'

const result = compileDeckSpecToNativeV1(spec)
if (!result.ok) throw new Error(result.issues.map(({ message }) => message).join('; '))`}
      />
      <p><code>@injoffice/slides</code> authors a <code>DeckSpec</code>. <code>pptx-authored</code> compiles it to native v1. <code>pptx-render</code> emits a renderer-neutral paint tree. Univer Slides is a separate, thinner OSS demo and is not this pipeline.</p>
    </Article>
  )
}
