import type { DemoDefinition, DemoFormat, DemoTask } from './demoRegistry'

export type ShowcaseFilters = {
  query: string
  task: DemoTask | 'All tasks'
  format: DemoFormat | 'All formats'
}

function searchableText(demo: DemoDefinition): string {
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
  ].join(' ').toLocaleLowerCase()
}

export function filterShowcaseDemos(demos: DemoDefinition[], filters: ShowcaseFilters): DemoDefinition[] {
  const terms = filters.query.trim().toLocaleLowerCase().split(/\s+/).filter(Boolean)

  return demos.filter((demo) => {
    if (filters.task !== 'All tasks' && !demo.tasks.includes(filters.task)) return false
    if (filters.format !== 'All formats' && !demo.formats.includes(filters.format)) return false
    if (terms.length === 0) return true

    const haystack = searchableText(demo)
    return terms.every((term) => haystack.includes(term))
  })
}
