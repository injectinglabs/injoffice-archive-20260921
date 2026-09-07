import { parseSurface, surfaceHref, type Surface } from './route'

export const SCROLL_SPY_PROBE_RATIO = 0.22

export function surfaceSectionId(surface: Surface): string {
  return `demo-${surface}`
}

export type SectionMetrics = {
  surface: Surface
  top: number
  bottom: number
}

/** Last section whose top has crossed the probe line, or the first section if none have. */
export function pickActiveSurface(sections: readonly SectionMetrics[], probeY: number): Surface | undefined {
  if (sections.length === 0) return undefined
  let active: Surface | undefined
  for (const section of sections) {
    if (section.top <= probeY) active = section.surface
  }
  return active ?? sections[0]?.surface
}

/** Keep query strings when the surface does not change; otherwise use a clean surface hash. */
export function hashForSurface(surface: Surface, currentHash: string): string {
  if (parseSurface(currentHash) === surface) return currentHash || surfaceHref(surface)
  return surfaceHref(surface)
}

export function prefersReducedMotion(media: Pick<MediaQueryList, 'matches'> | null | undefined = typeof window === 'undefined' ? undefined : window.matchMedia('(prefers-reduced-motion: reduce)')): boolean {
  return media?.matches === true
}

export function viewportProbeY(innerHeight = typeof window === 'undefined' ? 0 : window.innerHeight): number {
  return Math.max(88, Math.round(innerHeight * SCROLL_SPY_PROBE_RATIO))
}

export function scrollParent(element: HTMLElement): HTMLElement | 'window' {
  let current = element.parentElement
  while (current && current !== document.body && current !== document.documentElement) {
    const style = getComputedStyle(current)
    if ((style.overflowY === 'auto' || style.overflowY === 'scroll') && current.scrollHeight > current.clientHeight + 1) {
      return current
    }
    current = current.parentElement
  }
  return 'window'
}

export function scrollElementIntoView(element: HTMLElement, behavior: ScrollBehavior): void {
  const parent = scrollParent(element)
  if (parent === 'window') {
    const top = window.scrollY + element.getBoundingClientRect().top
    window.scrollTo({ top: Math.max(0, top), behavior })
    return
  }
  const nextTop = parent.scrollTop + element.getBoundingClientRect().top - parent.getBoundingClientRect().top
  parent.scrollTo({ top: Math.max(0, nextTop), behavior })
}
