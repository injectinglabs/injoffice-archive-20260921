# @injoffice/pptx-authored

Deterministically compiles InjOffice `DeckSpec` or `WireDeck` authoring data
into a validated, authored `pptx-native/v1` deck. The native deck is the input
to `@injoffice/pptx-render`; this package never measures HTML/CSS, imports a
browser renderer, takes screenshots, or rasterizes authored content.

Compilation is atomic. Unknown fields, inherited text styling, invalid or
fractional Office units, stale authoring keys, zero-extent connectors, and any
other unmodeled wire semantics return a refusal result with no partial deck.
Assets are explicit: the current WireDeck contract carries no assets, so the
compiled native manifest is always an explicit empty array.

```ts
import { compileDeckSpecToNativeV1 } from '@injoffice/pptx-authored'

const result = compileDeckSpecToNativeV1(spec)
if (!result.ok) throw new Error(result.issues.map(({ message }) => message).join('; '))
const nativeDeck = result.deck
```

The React `DeckView` and Konva `DeckCanvasView` exports remain temporarily for
legacy editing/preview consumers only. They are deprecated and are not Office
rendering authorities.
