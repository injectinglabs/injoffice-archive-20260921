import type { ComponentType, LazyExoticComponent } from 'react'
import { preloadableLazy, preloadableLazyNamed } from './preloadableLazy'
import type { Surface } from './route'

export type DemoRuntime = 'Browser' | 'Browser + local sidecar' | 'Local sidecar' | 'Node host'
export type DemoGroup = 'Create and edit' | 'Workbook tools' | 'Production pipeline'

export type DemoDefinition = {
  surface: Exclude<Surface, 'overview'>
  title: string
  navTitle: string
  packageName: string
  description: string
  runtime: DemoRuntime
  group: DemoGroup
  glyph: string
  accent: 'blue' | 'mint' | 'coral' | 'violet' | 'yellow'
  component: LazyExoticComponent<ComponentType>
  preload: () => Promise<void>
}

const sheets = preloadableLazy(() => import('./pages/SheetsPage'))
const docs = preloadableLazy(() => import('./pages/DocsPage'))
const slides = preloadableLazy(() => import('./pages/SlidesPage'))
const pdf = preloadableLazy(() => import('./pages/PdfPage'))
const charts = preloadableLazy(() => import('./pages/ChartsPage'))
const pivots = preloadableLazy(() => import('./pages/PivotsPage'))
const shapes = preloadableLazy(() => import('./pages/ShapesPage'))
const connectors = preloadableLazy(() => import('./pages/ConnectorsPage'))
const formulas = preloadableLazy(() => import('./pages/FormulasPage'))
const collab = preloadableLazyNamed(() => import('./collabPage'), 'CollabDemo')
const history = preloadableLazy(() => import('./pages/HistoryPage'))
const fontMetrics = preloadableLazy(() => import('./pages/FontMetricsPage'))
const pptxAuthored = preloadableLazy(() => import('./pages/PptxAuthoredPage'))
const pptxNative = preloadableLazy(() => import('./pages/PptxNativePage'))
const pptxRender = preloadableLazy(() => import('./pages/PptxRenderPage'))

export const DEMOS: DemoDefinition[] = [
  { surface: 'sheets', title: 'Spreadsheets', navTitle: 'Spreadsheets', packageName: '@injoffice/sheets', description: 'Edit a live workbook or run a browser-local native XLSX mutation, exact-byte readback, verification, and download.', runtime: 'Browser', group: 'Create and edit', glyph: 'S', accent: 'mint', ...sheets },
  { surface: 'docs', title: 'Documents', navTitle: 'Documents', packageName: '@injoffice/docs', description: 'Edit one guarded DOCX text run in browser-local bytes, reopen the exact output, verify preservation evidence, and download it.', runtime: 'Browser', group: 'Create and edit', glyph: 'D', accent: 'blue', ...docs },
  { surface: 'slides', title: 'Presentations', navTitle: 'Presentations', packageName: '@injoffice/slides', description: 'Turn an outline into an editable deck with themes, layout checks, transitions, and a live canvas.', runtime: 'Browser', group: 'Create and edit', glyph: 'P', accent: 'coral', ...slides },
  { surface: 'pdf', title: 'PDF workbench', navTitle: 'PDF', packageName: '@injoffice/pdf', description: 'Open a PDF and apply browser-safe page operations, markup, drawings, forms, and stamps, or download the bytes. Node PDFium text, image, OCR, and redaction run through the local Vite host during npm run dev.', runtime: 'Browser', group: 'Create and edit', glyph: 'A', accent: 'violet', ...pdf },
  { surface: 'charts', title: 'Charts', navTitle: 'Charts', packageName: '@injoffice/charts', description: 'Edit cell-range data and render the selected one of 30 chart types through the real ECharts adapter, with native wire and SVG export.', runtime: 'Browser', group: 'Workbook tools', glyph: 'C', accent: 'blue', ...charts },
  { surface: 'pivots', title: 'Pivot tables', navTitle: 'Pivot tables', packageName: '@injoffice/pivots', description: 'Group, filter, aggregate, and bake deterministic pivot results from a plain values grid.', runtime: 'Browser', group: 'Workbook tools', glyph: 'Σ', accent: 'mint', ...pivots },
  { surface: 'shapes', title: 'Shapes', navTitle: 'Shapes', packageName: '@injoffice/shapes', description: 'Browse 121 native shape identifiers through the package renderer, with distinct, approximate, and generic previews labeled.', runtime: 'Browser', group: 'Workbook tools', glyph: '◇', accent: 'coral', ...shapes },
  { surface: 'connectors', title: 'Data connectors', navTitle: 'Data connectors', packageName: '@injoffice/connectors', description: 'Normalize JSON and CSV into bound ranges while the host keeps credentials and network policy.', runtime: 'Browser', group: 'Workbook tools', glyph: '↳', accent: 'yellow', ...connectors },
  { surface: 'formulas', title: 'Formula coverage', navTitle: 'Formulas', packageName: '@injoffice/formulas', description: 'Search the audited business-function set and inspect valid formulas used in real engine qualification.', runtime: 'Browser', group: 'Workbook tools', glyph: 'ƒ', accent: 'violet', ...formulas },
  { surface: 'collab', title: 'Collaboration', navTitle: 'Collaboration', packageName: '@injoffice/collab', description: 'Edit two independent Univer sheets, a ProseMirror document, a DeckSpec, or a PDF annotator side by side. InjOffice synchronizes operations and presence in the browser; HTTP+SSE mode exercises the same rooms through the sidecar.', runtime: 'Browser', group: 'Production pipeline', glyph: '◎', accent: 'mint', ...collab },
  { surface: 'history', title: 'History and diffs', navTitle: 'History and diffs', packageName: '@injoffice/history', description: 'Explain document changes as structured cell and text diffs instead of opaque file versions.', runtime: 'Browser', group: 'Production pipeline', glyph: '↶', accent: 'yellow', ...history },
  { surface: 'font-metrics', title: 'Typography and layout', navTitle: 'Typography', packageName: '@injoffice/font-metrics', description: 'Inspect browser-safe layout contracts and the explicit Node boundary for font resolution, shaping, and bidi.', runtime: 'Browser', group: 'Production pipeline', glyph: 'Aa', accent: 'blue', ...fontMetrics },
  { surface: 'pptx-authored', title: 'PPTX authoring', navTitle: 'PPTX authoring', packageName: '@injoffice/pptx-authored', description: 'Compile a DeckSpec into strict native presentation objects with stable authored identities.', runtime: 'Browser', group: 'Production pipeline', glyph: '1', accent: 'coral', ...pptxAuthored },
  { surface: 'pptx-native', title: 'Native PPTX model', navTitle: 'Native PPTX', packageName: '@injoffice/pptx-native', description: 'Edit guarded text or one of four exact AutoShapes in browser-local PPTX bytes, reopen and verify the exact output, then download it.', runtime: 'Browser', group: 'Production pipeline', glyph: '{}', accent: 'violet', ...pptxNative },
  { surface: 'pptx-render', title: 'PPTX renderer', navTitle: 'PPTX renderer', packageName: '@injoffice/pptx-render', description: 'Compile native slides into deterministic render trees and renderer-independent paint commands.', runtime: 'Browser', group: 'Production pipeline', glyph: '▻', accent: 'yellow', ...pptxRender },
]

export const DEMO_BY_SURFACE = new Map(DEMOS.map((demo) => [demo.surface, demo]))
export const DEMO_GROUPS: DemoGroup[] = ['Create and edit', 'Workbook tools', 'Production pipeline']

export function preloadDemo(surface: Surface): Promise<void> {
  if (surface === 'overview') return Promise.resolve()
  return DEMO_BY_SURFACE.get(surface)?.preload() ?? Promise.resolve()
}

export function preloadDemoOnIntent(surface: Surface): void {
  void preloadDemo(surface).catch(() => undefined)
}
