import type { DeckSpec, SlideKind, SlideSpec } from './types'

// Deterministic (no LLM) layout QC for DeckSpec. DeckView.tsx is a fixed
// HTML/flex template per SlideKind with no absolute x/y/w/h — CSS
// container-query units (cqw) sized relative to the slide's own width,
// flowing top-to-bottom, clipped by `overflow:hidden` on the 6cqw-inset
// content box (SlideView).
//
// Because there is no real box geometry to compare, this estimates it: a
// character-width text-metrics heuristic (proportional sans-serif average
// glyph width ≈ 0.52× font-size) turns each text field's length + the
// template's known max-width into an estimated wrapped line count, and sums
// per-kind stacked element heights (title/eyebrow/subtitle/body/bullets, plus
// each kind's fixed gaps) against the 44.25cqw content-height budget that
// falls out of the 16:9 stage minus its 6cqw margins. That sum crossing the
// budget is this model's "out of bounds" — SlideView clips it, so overflow
// really does mean lost content, not just visual crowding. A genuine
// *overlap* check (two independent boxes intersecting) isn't meaningful for
// a single-column flow layout — the one real overlap risk here is stacked
// content growing tall enough to run into a kind's FIXED-position decoration
// (the accent bar under a title/closing slide, the page-number chip in a
// slide's bottom-right corner), so that risk is checked explicitly instead of
// faked as a generic AABB test. All thresholds are heuristic estimates, not
// pixel-exact — this is a coarse "is this dangerously long" filter, not a
// layout engine.
//
// When the OOXML canvas model lands (real x/y/w/h shapes — see
// go/pptxpatch), a second, exact overlap/overflow checker over THAT geometry
// can use AABB intersection; this checker stays as DeckSpec's own
// pre-compile QC pass.

export type QCIssueKind = 'overflow' | 'overlap' | 'out-of-bounds'

export interface QCIssue {
  slideId: string
  slideIndex: number
  kind: QCIssueKind
  /** Human-readable, meant to be fed straight back to a revise/fix prompt. */
  message: string
}

// ---- Stage geometry, all in "cqw" (% of slide width) — mirrors DeckView.tsx ----

const STAGE_W = 100
const STAGE_H = (STAGE_W * 9) / 16 // 16:9 aspect, height in the same cqw unit
const MARGIN = 6 // SlideView's inset on all 4 sides
const CONTENT_W = STAGE_W - 2 * MARGIN
const CONTENT_H = STAGE_H - 2 * MARGIN // ≈ 44.25

/** Reserved bottom strip: the "N / total" page-number chip every non-title
 *  kind renders at right/bottom ~2.4cqw — content stacking into this band
 *  risks visually colliding with it. */
const PAGE_NUMBER_RESERVE_H = 4
/** Reserved bottom strip for the accent bar under title/closing slides. */
const ACCENT_BAR_RESERVE_H = 3

/** Average glyph width as a fraction of font-size, for a proportional
 *  display/body sans stack — a standard typographic rule of thumb (not
 *  measured per font; this is an estimate, documented as such throughout). */
const AVG_CHAR_WIDTH = 0.52

/** Strip the bold/italic emphasis markers DeckView's renderInline consumes —
 *  they don't render as glyphs, so they shouldn't count toward width. */
function plainLength(text: string): number {
  return text.replace(/\*\*([^*]+)\*\*/g, '$1').replace(/\*([^*]+)\*/g, '$1').length
}

/** Estimated wrapped line count for one paragraph of text at fontSizeCqw,
 *  inside a box boxWidthCqw wide. Explicit '\n' segments (pre-wrap fields
 *  like body/notes) each wrap independently. Empty text = 0 lines. */
export function estimateWrappedLines(text: string, boxWidthCqw: number, fontSizeCqw: number): number {
  if (!text) return 0
  const charsPerLine = Math.max(1, Math.floor(boxWidthCqw / (fontSizeCqw * AVG_CHAR_WIDTH)))
  let lines = 0
  for (const segment of text.split('\n')) {
    const len = plainLength(segment)
    lines += len === 0 ? 1 : Math.ceil(len / charsPerLine)
  }
  return lines
}

/** Estimated rendered height (cqw) of one paragraph: lines × fontSize × lineHeight. */
function textBlockHeight(text: string, boxWidthCqw: number, fontSizeCqw: number, lineHeight: number): number {
  return estimateWrappedLines(text, boxWidthCqw, fontSizeCqw) * fontSizeCqw * lineHeight
}

/** Bulleted list height: each item wraps independently within boxWidthCqw
 *  (minus the list's left padding), items separated by the template's gap. */
function bulletsHeight(items: string[] | undefined, boxWidthCqw: number): number {
  const list = items ?? []
  if (list.length === 0) return 0
  const fontSizeCqw = 2.6
  const gap = 1.5
  const innerW = boxWidthCqw - 2.4 // ul padding-left
  const linesTotal = list.reduce((sum, item) => sum + Math.max(1, estimateWrappedLines(item, innerW, fontSizeCqw)), 0)
  return linesTotal * fontSizeCqw * 1.55 + gap * (list.length - 1)
}

function eyebrowHeight(eyebrow: string | undefined): number {
  return eyebrow ? 1.4 * 1.2 + 1.6 : 0 // font-size*lineHeight-ish + its own margin-bottom
}

interface KindBudget {
  /** Estimated stacked content height (cqw) for this slide. */
  contentH: number
  /** Bottom reserve this kind's fixed decoration needs (cqw), if any. */
  reserve: number
  reserveLabel: string
  /** Per-field notes for width-driven issues worth calling out individually. */
  fieldNotes: string[]
}

function budgetFor(slide: SlideSpec): KindBudget {
  const notes: string[] = []
  switch (slide.kind) {
    case 'title': {
      const titleW = CONTENT_W * 0.8
      const subW = CONTENT_W * 0.62
      const h =
        eyebrowHeight(slide.eyebrow) +
        textBlockHeight(slide.title ?? '', titleW, 7, 1.05) +
        (slide.subtitle ? textBlockHeight(slide.subtitle, subW, 2.6, 1.55) + 2.4 : 0) +
        (slide.body ? textBlockHeight(slide.body, subW, 2, 1.55) + 1.6 : 0)
      return { contentH: h, reserve: ACCENT_BAR_RESERVE_H, reserveLabel: 'the title bar accent', fieldNotes: notes }
    }
    case 'section': {
      const titleW = CONTENT_W * 0.78
      const subW = CONTENT_W * 0.64
      const h =
        eyebrowHeight(slide.eyebrow) +
        textBlockHeight(slide.title ?? '', titleW, 5.4, 1.05) +
        (slide.subtitle || slide.body ? textBlockHeight(slide.subtitle ?? slide.body ?? '', subW, 2.3, 1.55) + 2 : 0)
      return { contentH: h, reserve: PAGE_NUMBER_RESERVE_H, reserveLabel: 'the page-number chip', fieldNotes: notes }
    }
    case 'quote': {
      const quoteW = CONTENT_W * 0.8 - 5 // minus the callout box's own horizontal padding (5cqw)
      const h =
        8 + // callout vertical padding (4cqw top + 4cqw bottom)
        textBlockHeight(slide.quote ?? '', quoteW, 3.4, 1.4) +
        (slide.subtitle ? textBlockHeight(slide.subtitle, quoteW, 1.6, 1.4) + 2.2 : 0)
      return { contentH: h, reserve: PAGE_NUMBER_RESERVE_H, reserveLabel: 'the page-number chip', fieldNotes: notes }
    }
    case 'two-col': {
      const colW = (CONTENT_W - 5) / 2
      const leftH = bulletsHeight(slide.bullets, colW)
      const rightH = bulletsHeight(slide.bulletsRight, colW)
      if (Math.abs(leftH - rightH) > 12) {
        notes.push(
          `two-col columns are visually unbalanced (left ≈${leftH.toFixed(1)}cqw tall vs right ≈${rightH.toFixed(1)}cqw) — consider redistributing items`,
        )
      }
      const colH = Math.max(leftH, rightH)
      const h = eyebrowHeight(slide.eyebrow) + textBlockHeight(slide.title ?? '', CONTENT_W, 4, 1.05) + 3 + colH
      return { contentH: h, reserve: PAGE_NUMBER_RESERVE_H, reserveLabel: 'the page-number chip', fieldNotes: notes }
    }
    case 'closing': {
      const h =
        textBlockHeight(slide.title ?? 'Thank you', CONTENT_W, 5.6, 1.05) +
        (slide.subtitle ? textBlockHeight(slide.subtitle, CONTENT_W, 2.2, 1.55) + 2 : 0)
      return { contentH: h, reserve: ACCENT_BAR_RESERVE_H + 3.4, reserveLabel: 'the closing accent bar', fieldNotes: notes }
    }
    case 'bullets':
    default: {
      const bodyW = CONTENT_W * 0.8
      const h =
        eyebrowHeight(slide.eyebrow) +
        (slide.title ? textBlockHeight(slide.title, CONTENT_W, 4, 1.05) + 3 : 0) +
        (slide.body ? textBlockHeight(slide.body, bodyW, 2.1, 1.55) + 2.4 : 0) +
        bulletsHeight(slide.bullets, CONTENT_W)
      return { contentH: h, reserve: PAGE_NUMBER_RESERVE_H, reserveLabel: 'the page-number chip', fieldNotes: notes }
    }
  }
}

function fmt(n: number): string {
  return n.toFixed(1)
}

/** Audit one slide. Pure function — same input, same output, no rendering. */
export function auditSlide(slide: SlideSpec, slideIndex: number): QCIssue[] {
  const issues: QCIssue[] = []
  const budget = budgetFor(slide)
  const push = (kind: QCIssueKind, message: string) => issues.push({ slideId: slide.id, slideIndex, kind, message })

  for (const note of budget.fieldNotes) push('overlap', note)

  const available = CONTENT_H
  if (budget.contentH > available) {
    push(
      'out-of-bounds',
      `content is estimated at ≈${fmt(budget.contentH)}cqw tall, ${fmt(budget.contentH - available)}cqw past the ${fmt(available)}cqw slide content area — it will be clipped (shorten text, split into another slide, or reduce bullet count)`,
    )
  } else if (budget.reserve > 0 && budget.contentH > available - budget.reserve) {
    push(
      'overlap',
      `content is estimated at ≈${fmt(budget.contentH)}cqw tall, within ${fmt(available - budget.contentH)}cqw of ${budget.reserveLabel} — it may visually collide with it`,
    )
  }

  return issues
}

/** Audit every slide in a deck. Empty array = clean deck. */
export function auditDeck(deck: DeckSpec): QCIssue[] {
  const issues: QCIssue[] = []
  deck.slides.forEach((slide, i) => issues.push(...auditSlide(slide, i)))
  return issues
}

/** Format issues as text suitable for feeding back into an AI revise prompt. */
export function formatDeckAudit(issues: QCIssue[]): string {
  if (issues.length === 0) return '<deck-audit>Passed: no estimated overflow/overlap issues.</deck-audit>'
  const body = issues.map((iss) => `- Slide ${iss.slideIndex + 1} (${iss.slideId}) [${iss.kind}]: ${iss.message}`).join('\n')
  return `<deck-audit>Found ${issues.length} issue(s):\n${body}\n</deck-audit>`
}

export type { SlideKind }
