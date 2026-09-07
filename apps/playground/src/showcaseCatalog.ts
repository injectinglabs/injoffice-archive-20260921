import type { DemoDefinition, DemoFormat, DemoTask } from './demoRegistry'
import { AGENT_TOOLS, agentHref, surfaceHref, type AgentTool } from './route'

export type ShowcaseFilters = {
  query: string
  task: DemoTask | 'All tasks'
  format: DemoFormat | 'All formats'
}

export type ShowcaseItem = {
  key: string
  href: string
  surface: DemoDefinition['surface']
  accent: DemoDefinition['accent']
  glyph: string
  title: string
  description: string
  subtitle: string
  formats: DemoFormat[]
  tasks: DemoTask[]
  runtime: DemoDefinition['runtime']
  minutes: number
  searchText: string
}

const AGENT_SHOWCASE: Record<AgentTool, { title: string; description: string; format: DemoFormat; extra: string }> = {
  sheets: {
    title: 'Raise East forecast confidence',
    description: 'Inspect quarterly-plan.xlsx, set Forecast!D5 from Medium to High, and verify the replacement revision.',
    format: 'XLSX',
    extra: 'xlsx sheets cell forecast spreadsheet',
  },
  docs: {
    title: 'Replace a guarded document paragraph',
    description: 'Inspect the launch brief, replace one paragraph with the confirmed review date, and verify the output.',
    format: 'DOCX',
    extra: 'docx docs paragraph document word',
  },
  slides: {
    title: 'Update a presentation readiness metric',
    description: 'Inspect board-update.pptx, change the readiness metric from 86% to 91%, and verify the committed slide.',
    format: 'PPTX',
    extra: 'pptx slides deck metric presentation',
  },
  pdf: {
    title: 'Rotate a PDF review page',
    description: 'Inspect review-packet.pdf, rotate page two by 90°, and verify the page content is otherwise unchanged.',
    format: 'PDF',
    extra: 'pdf page rotate packet',
  },
}

function searchableText(parts: readonly string[]): string {
  return parts.join(' ').toLocaleLowerCase()
}

function demoSearchParts(demo: DemoDefinition): string[] {
  return [
    demo.title,
    demo.navTitle,
    demo.packageName,
    demo.description,
    demo.recipe.title,
    demo.recipe.outcome,
    demo.group,
    ...demo.tasks,
    ...demo.formats,
    ...(demo.keywords ?? []),
  ]
}

function matchesFilters(searchText: string, tasks: readonly DemoTask[], formats: readonly DemoFormat[], filters: ShowcaseFilters): boolean {
  if (filters.task !== 'All tasks' && !tasks.includes(filters.task)) return false
  if (filters.format !== 'All formats' && !formats.includes(filters.format)) return false
  const terms = filters.query.trim().toLocaleLowerCase().split(/\s+/).filter(Boolean)
  if (terms.length === 0) return true
  return terms.every((term) => searchText.includes(term))
}

export function showcaseItems(demos: readonly DemoDefinition[]): ShowcaseItem[] {
  return demos.flatMap((demo) => {
    if (demo.surface !== 'agent') {
      return [{
        key: demo.surface,
        href: demo.surface === 'sheets' ? `${surfaceHref('sheets')}?view=native` : surfaceHref(demo.surface),
        surface: demo.surface,
        accent: demo.accent,
        glyph: demo.glyph,
        title: demo.recipe.title,
        description: demo.recipe.outcome,
        subtitle: `${demo.title} · ${demo.tasks.join(' · ')}`,
        formats: demo.formats,
        tasks: demo.tasks,
        runtime: demo.runtime,
        minutes: demo.recipe.minutes,
        searchText: searchableText(demoSearchParts(demo)),
      }]
    }

    return AGENT_TOOLS.map((item) => {
      const extra = AGENT_SHOWCASE[item.tool]
      return {
        key: `agent-${item.tool}`,
        href: agentHref(item.tool),
        surface: demo.surface,
        accent: demo.accent,
        glyph: demo.glyph,
        title: extra.title,
        description: extra.description,
        subtitle: `${demo.title} · ${item.label} · ${demo.tasks.join(' · ')}`,
        formats: [extra.format],
        tasks: demo.tasks,
        runtime: demo.runtime,
        minutes: demo.recipe.minutes,
        searchText: searchableText([...demoSearchParts(demo), extra.title, extra.description, extra.format, item.label, extra.extra]),
      }
    })
  })
}

export function filterShowcaseDemos(demos: DemoDefinition[], filters: ShowcaseFilters): DemoDefinition[] {
  return demos.filter((demo) => matchesFilters(searchableText(demoSearchParts(demo)), demo.tasks, demo.formats, filters))
}

export function filterShowcaseItems(items: readonly ShowcaseItem[], filters: ShowcaseFilters): ShowcaseItem[] {
  return items.filter((item) => matchesFilters(item.searchText, item.tasks, item.formats, filters))
}
