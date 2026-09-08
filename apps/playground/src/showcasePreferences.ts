import { DEMO_FORMATS, DEMO_TASKS, type DemoFormat, type DemoTask } from './demoRegistry'

export type ShowcasePreferences = {
  query: string
  task: DemoTask | 'All tasks'
  format: DemoFormat | 'All formats'
}

const KEY = 'injoffice.showcase.filters.v1'
const defaults: ShowcasePreferences = { query: '', task: 'All tasks', format: 'All formats' }
type PreferenceStorage = Pick<Storage, 'getItem' | 'setItem'>

/** Tab-local only: returning from an example preserves the user's search. */
export function readShowcasePreferences(storage?: PreferenceStorage): ShowcasePreferences {
  try {
    const value = JSON.parse((storage ?? window.sessionStorage).getItem(KEY) ?? 'null') as Partial<ShowcasePreferences> | null
    return {
      query: typeof value?.query === 'string' ? value.query : defaults.query,
      task: DEMO_TASKS.includes(value?.task as DemoTask) ? value!.task! : defaults.task,
      format: DEMO_FORMATS.includes(value?.format as DemoFormat) ? value!.format! : defaults.format,
    }
  } catch {
    return { ...defaults }
  }
}

export function persistShowcasePreferences(value: ShowcasePreferences, storage?: PreferenceStorage): void {
  try { (storage ?? window.sessionStorage).setItem(KEY, JSON.stringify(value)) } catch {
    // A disabled/full session store must not prevent filtering the catalogue.
  }
}
