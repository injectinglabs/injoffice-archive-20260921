import type { DeckSpec, SlideSpec } from './types'

// deckFromOutline — the agent-facing shortcut: a markdown outline becomes a
// DeckSpec deterministically, so "make a deck about X" needs one text
// artifact, not JSON authoring. Rules:
//
//   # Title            → title slide (first #) / section slide (later #s)
//   ## Slide title     → bullets slide; its "- " lines become bullets
//   > quote — attrib   → quote slide
//   plain paragraph    → body text of the current slide
//   "Notes: …" line    → speaker notes of the current slide
//
// Anything unrecognized lands in body text — never dropped.

let seq = 0
function slideId(): string {
  return `slide-${(++seq).toString(36)}-${seq}`
}

/** Reset the id counter (tests). */
export function resetSlideIds(): void {
  seq = 0
}

export function deckFromOutline(title: string, outline: string): DeckSpec {
  const slides: SlideSpec[] = []
  let current = null as SlideSpec | null
  let sawTitle = false

  const push = (s: SlideSpec) => {
    slides.push(s)
    current = s
  }

  for (const raw of outline.split('\n')) {
    const line = raw.trimEnd()
    const t = line.trim()
    if (t === '') continue
    if (t.startsWith('# ')) {
      if (!sawTitle) {
        sawTitle = true
        push({ id: slideId(), kind: 'title', title: t.slice(2).trim() })
      } else {
        push({ id: slideId(), kind: 'section', title: t.slice(2).trim() })
      }
      continue
    }
    if (t.startsWith('## ')) {
      push({ id: slideId(), kind: 'bullets', title: t.slice(3).trim(), bullets: [] })
      continue
    }
    if (t.startsWith('> ')) {
      const q = t.slice(2).trim()
      const dash = q.lastIndexOf('—')
      push({
        id: slideId(),
        kind: 'quote',
        quote: dash > 0 ? q.slice(0, dash).trim() : q,
        subtitle: dash > 0 ? q.slice(dash + 1).trim() : undefined,
      })
      continue
    }
    if (/^notes:/i.test(t)) {
      if (current) current.notes = [current.notes, t.replace(/^notes:/i, '').trim()].filter(Boolean).join('\n')
      continue
    }
    if (t.startsWith('- ') || t.startsWith('* ')) {
      if (current && (current.kind === 'bullets' || current.kind === 'two-col')) {
        ;(current.bullets ??= []).push(t.slice(2).trim())
      } else {
        push({ id: slideId(), kind: 'bullets', bullets: [t.slice(2).trim()] })
      }
      continue
    }
    // Plain text: subtitle for a fresh title/section slide, body otherwise.
    if (current && (current.kind === 'title' || current.kind === 'section') && !current.subtitle) {
      current.subtitle = t
    } else if (current) {
      current.body = [current.body, t].filter(Boolean).join('\n')
    } else {
      push({ id: slideId(), kind: 'title', title, subtitle: t })
      sawTitle = true
    }
  }

  if (slides.length === 0) push({ id: slideId(), kind: 'title', title })
  const slug = title.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 40) || 'deck'
  return { id: `deck-${slug}-${slides.length}`, title, slides }
}
