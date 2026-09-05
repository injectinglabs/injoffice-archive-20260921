import { auditDeck, formatDeckAudit } from '../../../packages/slides/src/qc'
import type { AnimDirection, DeckSpec, SlideTransitionKind } from '../../../packages/slides/src/types'

export function playgroundDeckAudit(spec: DeckSpec) {
  const issues = auditDeck(spec)
  return { issues, report: formatDeckAudit(issues) }
}

export type SlideTransitionChoice = SlideTransitionKind | 'none'

export function updateSlideTransition(spec: DeckSpec, slideIndex: number, kind: SlideTransitionChoice, direction: AnimDirection = 'left'): DeckSpec {
  return {
    ...spec,
    slides: spec.slides.map((slide, index) => index !== slideIndex ? slide : {
      ...slide,
      transition: kind === 'none' ? undefined : kind === 'fade' ? { kind } : { kind, direction },
    }),
  }
}
