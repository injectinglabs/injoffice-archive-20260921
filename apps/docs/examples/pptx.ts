import { createPptxWasmClient } from '@injoffice/pptx-wasm'

/** Inspect native source identities before choosing an editable target. */
export async function inspectPresentation(original: Uint8Array) {
  const client = createPptxWasmClient()
  try {
    const deck = await client.extract(original)
    return {
      revision: deck.sourceRevision,
      textTargets: deck.slides.flatMap(slide => slide.elements
        .filter(element => element.kind === 'text' && element.source)
        .map(element => ({ slideId: slide.id, elementId: element.id, source: element.source }))),
    }
  } finally {
    client.terminate()
  }
}
