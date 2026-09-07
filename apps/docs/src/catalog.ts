export type DocRuntime = 'Browser' | 'Browser + sidecar' | 'Go sidecar' | 'Node host'
export type FileAuthority = 'Original OOXML bytes' | 'Original PDF bytes' | 'Plain JSON spec' | 'Host-owned' | 'Univer snapshot (UI only)'

export type DocId =
  | 'home'
  | 'showcase'
  | 'reference'
  | 'introduction'
  | 'installation'
  | 'concepts'
  | 'quickstart'
  | 'agent-workflows'
  | 'xlsx'
  | 'docx'
  | 'pptx'
  | 'sheets-editor'
  | 'charts'
  | 'pivots'
  | 'shapes'
  | 'connectors'
  | 'formulas'
  | 'collab'
  | 'history'
  | 'font-metrics'
  | 'pdf'
  | 'slides'
  | 'compare'

export type DocPage = {
  id: DocId
  href: string
  title: string
  navTitle: string
  group: string
  packageName?: string
  runtime?: DocRuntime
  authority?: FileAuthority
  description: string
  keywords: string[]
}

export const PAGES: DocPage[] = [
  {
    id: 'home',
    href: '#/guides',
    title: 'InjOffice documentation',
    navTitle: 'Home',
    group: 'Start',
    description: 'Native Office file engines with live examples and copy-paste usage.',
    keywords: ['home', 'docs'],
  },
  {
    id: 'introduction',
    href: '#/guides/introduction',
    title: 'What is InjOffice',
    navTitle: 'Introduction',
    group: 'Getting started',
    description: 'A toolkit for inspecting, editing, and verifying Office files without reconstructing what you do not understand.',
    keywords: ['intro', 'overview', 'office', 'ooxml'],
  },
  {
    id: 'installation',
    href: '#/guides/installation',
    title: 'Installation',
    navTitle: 'Installation',
    group: 'Getting started',
    description: 'Install individual packages. Nothing requires a hosted InjOffice service.',
    keywords: ['install', 'npm', 'workspace'],
  },
  {
    id: 'concepts',
    href: '#/guides/concepts',
    title: 'Basic concepts',
    navTitle: 'Concepts',
    group: 'Getting started',
    description: 'File authority, native JSON, mutations, the optional Univer shell, and fail-closed refusals.',
    keywords: ['authority', 'snapshot', 'mutation', 'univer', 'sidecar'],
  },
  {
    id: 'quickstart',
    href: '#/guides/quickstart',
    title: 'Quickstart',
    navTitle: 'Quickstart',
    group: 'Getting started',
    description: 'A working chart, a validated mutation batch, and a native extract in a few minutes.',
    keywords: ['quickstart', 'hello', 'first'],
  },
  {
    id: 'agent-workflows',
    href: '#/guides/agent-workflows',
    title: 'Agent change sets',
    navTitle: 'Agent workflows',
    group: 'Getting started',
    packageName: '@injoffice/agent-tools',
    runtime: 'Browser',
    authority: 'Host-owned',
    description: 'Provider-independent inspect, plan, preview, validate, approve, commit, and verify workflows for Office artifacts.',
    keywords: ['agent', 'ai', 'changeset', 'approval', 'verification', 'capabilities'],
  },
  {
    id: 'xlsx',
    href: '#/guides/xlsx',
    title: 'Native XLSX',
    navTitle: 'Native XLSX',
    group: 'Native files',
    packageName: '@injoffice/sheets',
    runtime: 'Browser + sidecar',
    authority: 'Original OOXML bytes',
    description: 'Extract a workbook to native JSON and apply a bounded mutation batch back onto the original package.',
    keywords: ['xlsx', 'extract', 'mutations', 'sheets'],
  },
  {
    id: 'docx',
    href: '#/guides/docx',
    title: 'Native DOCX',
    navTitle: 'Native DOCX',
    group: 'Native files',
    packageName: '@injoffice/docs',
    runtime: 'Browser + sidecar',
    authority: 'Original OOXML bytes',
    description: 'Decode a native document projection and replace one guarded text run in the original DOCX.',
    keywords: ['docx', 'document', 'run', 'word'],
  },
  {
    id: 'pptx',
    href: '#/guides/pptx',
    title: 'Native PPTX',
    navTitle: 'Native PPTX',
    group: 'Native files',
    packageName: '@injoffice/pptx-native',
    runtime: 'Browser + sidecar',
    authority: 'Original OOXML bytes',
    description: 'Extract a presentation, compile an authored deck, and apply bounded native mutations.',
    keywords: ['pptx', 'slides', 'presentation'],
  },
  {
    id: 'sheets-editor',
    href: '#/guides/sheets-editor',
    title: 'Sheets editor (Univer OSS)',
    navTitle: 'Sheets editor',
    group: 'Editor',
    packageName: '@injoffice/univer-sheets',
    runtime: 'Browser',
    authority: 'Univer snapshot (UI only)',
    description: 'Compose Apache-2.0 Univer Sheets as an optional grid. Univer is not the file authority.',
    keywords: ['univer', 'editor', 'preset', 'ribbon'],
  },
  {
    id: 'charts',
    href: '#/guides/charts',
    title: 'Charts',
    navTitle: 'Charts',
    group: 'Workbook tools',
    packageName: '@injoffice/charts',
    runtime: 'Browser',
    authority: 'Plain JSON spec',
    description: 'Renderer-neutral ChartSpec, data extraction, and ECharts options. Optional Univer float integration.',
    keywords: ['charts', 'echarts', 'column', 'series'],
  },
  {
    id: 'pivots',
    href: '#/guides/pivots',
    title: 'Pivot tables',
    navTitle: 'Pivot tables',
    group: 'Workbook tools',
    packageName: '@injoffice/pivots',
    runtime: 'Browser',
    authority: 'Plain JSON spec',
    description: 'Deterministic aggregation from a values grid. Optional Univer manager and native XLSX pivot parts.',
    keywords: ['pivot', 'aggregate', 'sum', 'group'],
  },
  {
    id: 'shapes',
    href: '#/guides/shapes',
    title: 'Shapes',
    navTitle: 'Shapes',
    group: 'Workbook tools',
    packageName: '@injoffice/shapes',
    runtime: 'Browser',
    authority: 'Plain JSON spec',
    description: 'ECMA-376 preset geometries as plain JSON, with SVG previews and optional Univer floats.',
    keywords: ['shapes', 'drawing', 'preset', 'geometry'],
  },
  {
    id: 'connectors',
    href: '#/guides/connectors',
    title: 'Data connectors',
    navTitle: 'Connectors',
    group: 'Workbook tools',
    packageName: '@injoffice/connectors',
    runtime: 'Browser',
    authority: 'Host-owned',
    description: 'Normalize JSON and CSV into bound ranges. Credentials and network policy stay with the host.',
    keywords: ['csv', 'json', 'connector', 'fetch'],
  },
  {
    id: 'formulas',
    href: '#/guides/formulas',
    title: 'Formula coverage',
    navTitle: 'Formulas',
    group: 'Workbook tools',
    packageName: '@injoffice/formulas',
    runtime: 'Browser',
    authority: 'Host-owned',
    description: 'An auditable function matrix. Calculation remains the selected workbook engine.',
    keywords: ['formula', 'xlookup', 'sum', 'audit'],
  },
  {
    id: 'collab',
    href: '#/guides/collaboration',
    title: 'Collaboration',
    navTitle: 'Collaboration',
    group: 'Production',
    packageName: '@injoffice/collab',
    runtime: 'Browser + sidecar',
    authority: 'Host-owned',
    description: 'Presence and ordered operations. The package includes no server; an optional Go hub is in-repo.',
    keywords: ['collab', 'presence', 'ops', 'sse'],
  },
  {
    id: 'history',
    href: '#/guides/history',
    title: 'History and diffs',
    navTitle: 'History',
    group: 'Production',
    packageName: '@injoffice/history',
    runtime: 'Browser',
    authority: 'Host-owned',
    description: 'Pure grid and text diffs. Version storage stays with the host.',
    keywords: ['diff', 'history', 'changes'],
  },
  {
    id: 'font-metrics',
    href: '#/guides/font-metrics',
    title: 'Typography and layout',
    navTitle: 'Font metrics',
    group: 'Production',
    packageName: '@injoffice/font-metrics',
    runtime: 'Browser',
    authority: 'Host-owned',
    description: 'Renderer-neutral native text-layout contract. Font file discovery is Node-only.',
    keywords: ['font', 'harfbuzz', 'shaping', 'layout'],
  },
  {
    id: 'pdf',
    href: '#/guides/pdf',
    title: 'PDF',
    navTitle: 'PDF',
    group: 'Production',
    packageName: '@injoffice/pdf',
    runtime: 'Browser',
    authority: 'Original PDF bytes',
    description: 'Inspect, rotate, and transform original PDF bytes. Editing APIs exist beyond the viewer demo.',
    keywords: ['pdf', 'viewer', 'rotate'],
  },
  {
    id: 'slides',
    href: '#/guides/slides',
    title: 'Slides and PPTX authoring',
    navTitle: 'Slides',
    group: 'Production',
    packageName: '@injoffice/slides',
    runtime: 'Browser',
    authority: 'Plain JSON spec',
    description: 'Compile a DeckSpec into native PPTX objects, then render a deterministic paint tree.',
    keywords: ['slides', 'deckspec', 'authored'],
  },
  {
    id: 'compare',
    href: '#/guides/univer',
    title: 'Univer OSS and Pro',
    navTitle: 'Univer OSS / Pro',
    group: 'Compare',
    description: 'How InjOffice uses Apache Univer Sheets, what Pro adds, and what this toolkit implements independently.',
    keywords: ['univer', 'pro', 'comparison', 'license'],
  },
  {
    id: 'showcase',
    href: '#/guides/showcase',
    title: 'Showcase',
    navTitle: 'Showcase',
    group: 'Explore',
    description: 'Live examples that import the same packages your application would.',
    keywords: ['showcase', 'demo', 'examples'],
  },
  {
    id: 'reference',
    href: '#/guides/reference',
    title: 'Package reference',
    navTitle: 'Packages',
    group: 'Explore',
    description: 'Every published TypeScript package and the Go modules behind native extract/apply.',
    keywords: ['packages', 'api', 'reference'],
  },
]

export const PAGE_BY_ID = new Map(PAGES.map((page) => [page.id, page]))

export const NAV_GROUPS = [
  'Getting started',
  'Native files',
  'Editor',
  'Workbook tools',
  'Production',
  'Compare',
  'Explore',
] as const

export function pageHref(id: DocId): string {
  return PAGE_BY_ID.get(id)?.href ?? '#/'
}
