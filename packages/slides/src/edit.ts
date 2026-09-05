import type { DeckSpec, SlideKind, SlideSpec } from './types'

// Pure deck-editing operations (S3). Every function returns a NEW spec —
// the host owns state and dirty tracking; nothing here mutates. Kept free
// of React so the operations unit-test directly and any host (panel UI,
// agent RPC, tests) drives the same code.

let seq = 0

/** Deterministic-enough slide id for editor-created slides. */
function newSlideId(spec: DeckSpec): string {
  const taken = new Set(spec.slides.map((s) => s.id))
  for (;;) {
    const id = `slide-new-${++seq}`
    if (!taken.has(id)) return id
  }
}

/** Replace one slide's fields (id preserved). */
export function updateSlide(spec: DeckSpec, index: number, patch: Partial<Omit<SlideSpec, 'id'>>): DeckSpec {
  if (index < 0 || index >= spec.slides.length) return spec
  const slides = spec.slides.slice()
  slides[index] = { ...slides[index], ...patch, id: slides[index].id }
  return { ...spec, slides }
}

/** Insert a new slide of `kind` after `index` (or at the end for -1). */
export function addSlide(spec: DeckSpec, index: number, kind: SlideKind): DeckSpec {
  const blank: SlideSpec = { id: newSlideId(spec), kind }
  switch (kind) {
    case 'title':
      blank.title = 'Title'
      break
    case 'section':
      blank.title = 'Section'
      break
    case 'quote':
      blank.quote = 'Quote'
      break
    case 'two-col':
      blank.title = 'Compare'
      blank.bullets = ['Left point']
      blank.bulletsRight = ['Right point']
      break
    case 'closing':
      blank.title = 'Thank you'
      break
    default:
      blank.title = 'New slide'
      blank.bullets = ['First point']
  }
  const at = index < 0 ? spec.slides.length : Math.min(index + 1, spec.slides.length)
  const slides = [...spec.slides.slice(0, at), blank, ...spec.slides.slice(at)]
  return { ...spec, slides }
}

/** Remove the slide at index. A deck always keeps at least one slide. */
export function removeSlide(spec: DeckSpec, index: number): DeckSpec {
  if (spec.slides.length <= 1 || index < 0 || index >= spec.slides.length) return spec
  const slides = spec.slides.slice()
  slides.splice(index, 1)
  return { ...spec, slides }
}

/** Move the slide at index by delta (clamped). */
export function moveSlide(spec: DeckSpec, index: number, delta: number): DeckSpec {
  const to = index + delta
  if (index < 0 || index >= spec.slides.length || to < 0 || to >= spec.slides.length) return spec
  const slides = spec.slides.slice()
  const [s] = slides.splice(index, 1)
  slides.splice(to, 0, s)
  return { ...spec, slides }
}

/** Bullets as one editable text block (one bullet per line) and back. */
export function bulletsToText(bullets: string[] | undefined): string {
  return (bullets ?? []).join('\n')
}

export function textToBullets(text: string): string[] | undefined {
  const items = text
    .split('\n')
    .map((l) => l.replace(/^[-*]\s+/, '').trim())
    .filter(Boolean)
  return items.length > 0 ? items : undefined
}
