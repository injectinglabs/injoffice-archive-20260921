import { lazy, Suspense } from 'react'
import { ChartsBench } from '../examples/ChartsBench'
import { CollabBench } from '../examples/CollabBench'
import { ConnectorsBench } from '../examples/ConnectorsBench'
import { FontMetricsBench } from '../examples/FontMetricsBench'
import { FormulasBench } from '../examples/FormulasBench'
import { HistoryBench } from '../examples/HistoryBench'
import { MutationsBench } from '../examples/MutationsBench'
import { PdfBench } from '../examples/PdfBench'
import { PivotsBench } from '../examples/PivotsBench'
import { ShapesBench } from '../examples/ShapesBench'
import { SlidesBench } from '../examples/SlidesBench'

const EditorBench = lazy(() => import('../examples/EditorBench').then((mod) => ({ default: mod.EditorBench })))

export function ShowcasePage() {
  return (
    <article className="article">
      <h1>Showcase</h1>
      <p className="lead">Each preview imports a real InjOffice package, the same way Univer’s showcase mounts a running editor next to the snippet.</p>
      <h2 id="charts">Charts</h2>
      <p><a href="#/guides/charts">Guide and copy-paste API →</a></p>
      <ChartsBench />
      <h2 id="pivots">Pivot tables</h2>
      <p><a href="#/guides/pivots">Guide and copy-paste API →</a></p>
      <PivotsBench />
      <h2 id="shapes">Shapes</h2>
      <p><a href="#/guides/shapes">Guide and copy-paste API →</a></p>
      <ShapesBench />
      <h2 id="mutations">Workbook mutations</h2>
      <p><a href="#/guides/xlsx">Guide and copy-paste API →</a></p>
      <MutationsBench />
      <h2 id="history">History diffs</h2>
      <p><a href="#/guides/history">Guide and copy-paste API →</a></p>
      <HistoryBench />
      <h2 id="connectors">Connectors</h2>
      <p><a href="#/guides/connectors">Guide and copy-paste API →</a></p>
      <ConnectorsBench />
      <h2 id="formulas">Formulas</h2>
      <p><a href="#/guides/formulas">Guide and copy-paste API →</a></p>
      <FormulasBench />
      <h2 id="pdf">PDF</h2>
      <p><a href="#/guides/pdf">Guide and copy-paste API →</a></p>
      <PdfBench />
      <h2 id="slides">Slides</h2>
      <p><a href="#/guides/slides">Guide and copy-paste API →</a></p>
      <SlidesBench />
      <h2 id="collab">Collaboration transforms</h2>
      <p><a href="#/guides/collaboration">Guide and copy-paste API →</a></p>
      <CollabBench />
      <h2 id="fonts">Font metrics</h2>
      <p><a href="#/guides/font-metrics">Guide and copy-paste API →</a></p>
      <FontMetricsBench />
      <h2 id="editor">Univer OSS editor</h2>
      <p><a href="#/guides/sheets-editor">Guide and copy-paste API →</a></p>
      <Suspense fallback={<p>Loading Univer OSS editor…</p>}>
        <EditorBench />
      </Suspense>
    </article>
  )
}
