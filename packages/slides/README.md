# @injoffice/slides

A presentation authoring model with a DOM-free `./authoring` entry, editing
helpers, quality checks, diagrams, and deprecated React/Konva preview surfaces.
Production Office rendering authority belongs to `@injoffice/pptx-authored`
and `@injoffice/pptx-render`, not these legacy views.

```bash
npm install @injoffice/slides react react-dom
```

```ts
import { auditDeck, compileDeckToWire, type DeckSpec } from '@injoffice/slides'

const deck: DeckSpec = {
  id: 'demo',
  title: 'Quarterly review',
  slides: [{ id: 'intro', kind: 'title', title: 'Quarterly review', subtitle: 'Q2' }],
}

const wire = compileDeckToWire(deck)
const issues = auditDeck(deck)
```

`DeckSpec` is the persistence contract. The package provides browser renderers but no storage, collaboration server, or PowerPoint writer; the repository's Go patch library maps wire decks into PPTX.
