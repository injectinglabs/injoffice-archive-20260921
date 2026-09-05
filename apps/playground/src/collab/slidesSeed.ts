import { BUILTIN_THEMES, deckFromOutline, resetSlideIds, type DeckSpec } from '@injoffice/slides'

export const COLLAB_SLIDES_OUTLINE = `# InjOffice slides
A live DeckSpec proof for hosts.
## What the engine owns
- Plain-JSON DeckSpec
- Built-in themes a host can replace
- Konva canvas from the same compile path as pptx
## What the host owns
- Chrome, routing, brand
- Save/export through Go, not a second OPC writer
> Specs are the contract — the renderer is a proof
# Close
Ship the engine. Keep the press marks.
`

export function seedCollabDeck(): DeckSpec {
  resetSlideIds()
  const next = deckFromOutline('InjOffice slides', COLLAB_SLIDES_OUTLINE)
  next.theme = BUILTIN_THEMES[0]!.id
  return next
}
