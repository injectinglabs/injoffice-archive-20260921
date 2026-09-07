import type { ComponentType, LazyExoticComponent } from 'react'
import { DEMO_RECIPES, type DemoRecipe } from './demoRecipes'
import { preloadableLazy, preloadableLazyNamed } from './preloadableLazy'
import type { Surface } from './route'

export type DemoRuntime = 'Browser' | 'Browser + local sidecar' | 'Local sidecar' | 'Node host'
export type DemoGroup = 'Create and edit' | 'Workbook tools' | 'Production pipeline'
export type DemoTask = 'Edit files' | 'Create content' | 'Analyze data' | 'Review changes' | 'Ship workflows'
export type DemoFormat = 'XLSX' | 'DOCX' | 'PPTX' | 'PDF' | 'Cross-format'

export type DemoDefinition = {
  surface: Exclude<Surface, 'overview'>
  title: string
  navTitle: string
  packageName: string
  description: string
  runtime: DemoRuntime
  group: DemoGroup
  tasks: DemoTask[]
  formats: DemoFormat[]
  keywords?: string[]
  glyph: string
  accent: 'blue' | 'mint' | 'coral' | 'violet' | 'yellow'
  recipe: DemoRecipe
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
const agent = preloadableLazy(() => import('./pages/AgentPage'))
const collab = preloadableLazyNamed(() => import('./collabPage'), 'CollabDemo')
const history = preloadableLazy(() => import('./pages/HistoryPage'))
const fontMetrics = preloadableLazy(() => import('./pages/FontMetricsPage'))
const pptxAuthored = preloadableLazy(() => import('./pages/PptxAuthoredPage'))
const pptxNative = preloadableLazy(() => import('./pages/PptxNativePage'))
const pptxRender = preloadableLazy(() => import('./pages/PptxRenderPage'))

export const DEMOS: DemoDefinition[] = [
  { surface: 'sheets', title: 'Spreadsheets', navTitle: 'Spreadsheets', packageName: '@injoffice/sheets', description: 'Edit a live workbook or run a browser-local native XLSX mutation, exact-byte readback, verification, and download.', runtime: 'Browser', group: 'Create and edit', tasks: ['Edit files', 'Analyze data', 'Review changes'], formats: ['XLSX'], keywords: ['workbook', 'cells', 'native mutation', 'download'], glyph: 'S', accent: 'mint', recipe: DEMO_RECIPES.sheets, ...sheets },
  { surface: 'docs', title: 'Documents', navTitle: 'Documents', packageName: '@injoffice/docs', description: 'Edit one guarded DOCX text run in browser-local bytes, reopen the exact output, verify preservation evidence, and download it.', runtime: 'Browser', group: 'Create and edit', tasks: ['Edit files', 'Review changes'], formats: ['DOCX'], keywords: ['word', 'text', 'native mutation', 'download'], glyph: 'D', accent: 'blue', recipe: DEMO_RECIPES.docs, ...docs },
  { surface: 'slides', title: 'Presentations', navTitle: 'Presentations', packageName: '@injoffice/slides', description: 'Turn an outline into an editable deck with themes, layout checks, transitions, and a live canvas.', runtime: 'Browser', group: 'Create and edit', tasks: ['Create content', 'Edit files'], formats: ['PPTX'], keywords: ['deck', 'outline', 'themes', 'transitions'], glyph: 'P', accent: 'coral', recipe: DEMO_RECIPES.slides, ...slides },
  { surface: 'pdf', title: 'PDF workbench', navTitle: 'PDF', packageName: '@injoffice/pdf', description: 'Open a PDF and apply browser-safe page operations, markup, drawings, forms, and stamps, or download the bytes. Node PDFium text, image, OCR, and redaction run through the local Vite host during npm run dev.', runtime: 'Browser', group: 'Create and edit', tasks: ['Edit files', 'Review changes', 'Ship workflows'], formats: ['PDF'], keywords: ['annotate', 'markup', 'forms', 'OCR', 'redact', 'stamp'], glyph: 'A', accent: 'violet', recipe: DEMO_RECIPES.pdf, ...pdf },
  { surface: 'charts', title: 'Charts', navTitle: 'Charts', packageName: '@injoffice/charts', description: 'Edit cell-range data and render the selected one of 30 chart types through the real ECharts adapter, with native wire and SVG export.', runtime: 'Browser', group: 'Workbook tools', tasks: ['Analyze data', 'Create content'], formats: ['XLSX'], keywords: ['visualize', 'ECharts', 'SVG', 'range'], glyph: 'C', accent: 'blue', recipe: DEMO_RECIPES.charts, ...charts },
  { surface: 'pivots', title: 'Pivot tables', navTitle: 'Pivot tables', packageName: '@injoffice/pivots', description: 'Group, filter, aggregate, and bake deterministic pivot results from a plain values grid.', runtime: 'Browser', group: 'Workbook tools', tasks: ['Analyze data'], formats: ['XLSX'], keywords: ['aggregate', 'group', 'filter', 'table'], glyph: 'Σ', accent: 'mint', recipe: DEMO_RECIPES.pivots, ...pivots },
  { surface: 'shapes', title: 'Shapes', navTitle: 'Shapes', packageName: '@injoffice/shapes', description: 'Browse 121 native shape identifiers through the package renderer, with distinct, approximate, and generic previews labeled.', runtime: 'Browser', group: 'Workbook tools', tasks: ['Create content', 'Review changes'], formats: ['XLSX', 'PPTX'], keywords: ['drawing', 'AutoShape', 'renderer'], glyph: '◇', accent: 'coral', recipe: DEMO_RECIPES.shapes, ...shapes },
  { surface: 'connectors', title: 'Data connectors', navTitle: 'Data connectors', packageName: '@injoffice/connectors', description: 'Normalize JSON and CSV into bound ranges while the host keeps credentials and network policy.', runtime: 'Browser', group: 'Workbook tools', tasks: ['Analyze data', 'Ship workflows'], formats: ['XLSX'], keywords: ['JSON', 'CSV', 'import', 'bound ranges'], glyph: '↳', accent: 'yellow', recipe: DEMO_RECIPES.connectors, ...connectors },
  { surface: 'formulas', title: 'Formula coverage', navTitle: 'Formulas', packageName: '@injoffice/formulas', description: 'Search the audited business-function set and inspect valid formulas used in real engine qualification.', runtime: 'Browser', group: 'Workbook tools', tasks: ['Analyze data', 'Review changes'], formats: ['XLSX'], keywords: ['functions', 'calculation', 'audit'], glyph: 'ƒ', accent: 'violet', recipe: DEMO_RECIPES.formulas, ...formulas },
  { surface: 'agent', title: 'Agent change sets', navTitle: 'Agent workflows', packageName: '@injoffice/agent-tools', description: 'Let any model inspect, propose, preview, validate, and verify bounded changes while the host controls approval and commit.', runtime: 'Browser', group: 'Production pipeline', tasks: ['Edit files', 'Review changes', 'Ship workflows'], formats: ['Cross-format'], keywords: ['AI', 'agent', 'approval', 'change set', 'verification'], glyph: 'AI', accent: 'blue', recipe: DEMO_RECIPES.agent, ...agent },
  { surface: 'collab', title: 'Collaboration', navTitle: 'Collaboration', packageName: '@injoffice/collab', description: 'Edit two independent Univer sheets, a ProseMirror document, a DeckSpec, or a PDF annotator side by side. InjOffice synchronizes operations and presence in the browser; HTTP+SSE mode exercises the same rooms through the sidecar.', runtime: 'Browser', group: 'Production pipeline', tasks: ['Edit files', 'Ship workflows'], formats: ['Cross-format'], keywords: ['sync', 'presence', 'rooms', 'SSE'], glyph: '◎', accent: 'mint', recipe: DEMO_RECIPES.collab, ...collab },
  { surface: 'history', title: 'History and diffs', navTitle: 'History and diffs', packageName: '@injoffice/history', description: 'Explain document changes as structured cell and text diffs instead of opaque file versions.', runtime: 'Browser', group: 'Production pipeline', tasks: ['Review changes', 'Ship workflows'], formats: ['XLSX', 'DOCX'], keywords: ['version', 'audit', 'diff'], glyph: '↶', accent: 'yellow', recipe: DEMO_RECIPES.history, ...history },
  { surface: 'font-metrics', title: 'Typography and layout', navTitle: 'Typography', packageName: '@injoffice/font-metrics', description: 'Inspect browser-safe layout contracts and the explicit Node boundary for font resolution, shaping, and bidi.', runtime: 'Browser', group: 'Production pipeline', tasks: ['Review changes', 'Ship workflows'], formats: ['DOCX', 'PPTX', 'PDF'], keywords: ['fonts', 'shaping', 'bidi', 'layout'], glyph: 'Aa', accent: 'blue', recipe: DEMO_RECIPES['font-metrics'], ...fontMetrics },
  { surface: 'pptx-authored', title: 'PPTX authoring', navTitle: 'PPTX authoring', packageName: '@injoffice/pptx-authored', description: 'Compile a DeckSpec into strict native presentation objects with stable authored identities.', runtime: 'Browser', group: 'Production pipeline', tasks: ['Create content', 'Ship workflows'], formats: ['PPTX'], keywords: ['deck', 'compile', 'authoring'], glyph: '1', accent: 'coral', recipe: DEMO_RECIPES['pptx-authored'], ...pptxAuthored },
  { surface: 'pptx-native', title: 'Native PPTX model', navTitle: 'Native PPTX', packageName: '@injoffice/pptx-native', description: 'Edit guarded text or one of four exact AutoShapes in browser-local PPTX bytes, reopen and verify the exact output, then download it.', runtime: 'Browser', group: 'Production pipeline', tasks: ['Edit files', 'Review changes'], formats: ['PPTX'], keywords: ['text', 'AutoShape', 'native mutation', 'download'], glyph: '{}', accent: 'violet', recipe: DEMO_RECIPES['pptx-native'], ...pptxNative },
  { surface: 'pptx-render', title: 'PPTX renderer', navTitle: 'PPTX renderer', packageName: '@injoffice/pptx-render', description: 'Compile native slides into deterministic render trees and renderer-independent paint commands.', runtime: 'Browser', group: 'Production pipeline', tasks: ['Create content', 'Review changes', 'Ship workflows'], formats: ['PPTX'], keywords: ['paint', 'render tree', 'slides'], glyph: '▻', accent: 'yellow', recipe: DEMO_RECIPES['pptx-render'], ...pptxRender },
]

export const DEMO_BY_SURFACE = new Map(DEMOS.map((demo) => [demo.surface, demo]))
export const DEMO_GROUPS: DemoGroup[] = ['Create and edit', 'Workbook tools', 'Production pipeline']
export const DEMO_TASKS: DemoTask[] = ['Edit files', 'Create content', 'Analyze data', 'Review changes', 'Ship workflows']
export const DEMO_FORMATS: DemoFormat[] = ['XLSX', 'DOCX', 'PPTX', 'PDF', 'Cross-format']

export function preloadDemo(surface: Surface): Promise<void> {
  if (surface === 'overview') return Promise.resolve()
  return DEMO_BY_SURFACE.get(surface)?.preload() ?? Promise.resolve()
}

export function preloadDemoOnIntent(surface: Surface): void {
  void preloadDemo(surface).catch(() => undefined)
}
