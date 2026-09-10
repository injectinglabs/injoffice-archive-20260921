import { compileDeckSpecToNativeV1 } from '@injoffice/pptx-authored'
import { auditDeck, type DeckSpec } from '@injoffice/slides/authoring'

export function createQuarterlyReview() {
  const spec: DeckSpec = {
    id: 'quarterly-review', title: 'Quarterly review',
    slides: [{ id: 'intro', kind: 'title', title: 'Quarterly review', subtitle: 'Q3 operating review' }],
  }
  const audit = auditDeck(spec)
  const compiled = compileDeckSpecToNativeV1(spec)
  if (!compiled.ok) throw new Error(JSON.stringify(compiled.issues))
  // A native model is not a .pptx archive. Saving requires the native writer.
  return { spec, audit, nativeDeck: compiled.deck }
}
