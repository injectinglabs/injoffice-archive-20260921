import { ChartsBench } from '../examples/ChartsBench'
import { CodeBlock } from '../components/CodeBlock'

export function HomePage() {
  return (
    <article className="article">
      <p className="kicker"><span className="badge badge--pass">Apache-2.0</span><span className="badge">No @univerjs-pro</span></p>
      <div className="home-hero">
        <div>
          <h1>Edit the original Office file. Leave everything else byte-identical.</h1>
          <p className="lead">InjOffice is a set of TypeScript and Go engines for XLSX, DOCX, PPTX, and PDF. Each guide below has the install command, the code you actually run, and a live example that imports the same package.</p>
          <div className="actions">
            <a className="btn btn-primary" href="#/guides/quickstart">Quickstart</a>
            <a className="btn btn-ghost" href="#/guides/showcase">Open the showcase</a>
          </div>
        </div>
        <ChartsBench />
      </div>
      <h2 id="how-to-read">How these docs are organized</h2>
      <p>The layout follows the same working order as Univer’s developer docs: a concept, an install line, a copy-paste snippet, then a running preview. The difference is the subject. Univer documents an editor snapshot. InjOffice documents a file you can put back on disk.</p>
      <CodeBlock
        language="ts"
        title="Usage"
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
      <div className="card-grid">
        <a className="card" href="#/guides/xlsx"><small>@injoffice/sheets</small><strong>Native XLSX</strong><p>Extract and patch the original workbook.</p></a>
        <a className="card" href="#/guides/sheets-editor"><small>@injoffice/univer-sheets</small><strong>Sheets editor</strong><p>Optional Apache Univer grid. Not file authority.</p></a>
        <a className="card" href="#/guides/univer"><small>Compare</small><strong>Univer OSS / Pro</strong><p>What we reuse, what Pro licenses, what we implement.</p></a>
        <a className="card" href="#/guides/collaboration"><small>@injoffice/collab</small><strong>Collaboration</strong><p>Presence and ops without a commercial server.</p></a>
      </div>
    </article>
  )
}
