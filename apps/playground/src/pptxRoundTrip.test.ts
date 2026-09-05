import { describe, expect, it } from 'vitest'
import type { NativePptxDeck } from '@injoffice/pptx-native'
import {
  buildPptxMutationEvidence,
  buildPptxShapeMutation,
  buildPptxTextMutation,
  editablePptxShapeTargets,
  editablePptxTargets,
  editablePptxTextTargets,
  findPptxTextTarget,
  firstPptxRunText,
  pptxDownloadName,
  pptxTargetKey,
  untouchedPptxElementInventory,
  verifyPptxRoundTrip,
} from './pptxRoundTrip'

const REVISION = `rev-${'a'.repeat(64)}`
const FINGERPRINT = 'b'.repeat(64)

function fixtureDeck(): NativePptxDeck {
  return {
    contractVersion: 'pptx-native/v1',
    documentId: 'deck-fixture',
    origin: 'parsed',
    sourceRevision: REVISION,
    size: { cx: 12_192_000, cy: 6_858_000 },
    assets: [],
    slides: [{
      id: 'slide-one',
      provenance: 'parsed',
      elements: [{
        kind: 'group',
        id: 'group-one',
        provenance: 'parsed',
        transform: { x: 0, y: 0, cx: 10, cy: 10 },
        childTransform: { x: 0, y: 0, cx: 10, cy: 10 },
        source: { partName: 'ppt/slides/slide1.xml', objectId: 'group-1', fingerprintSha256: 'c'.repeat(64) },
        passthrough: [],
        compatibility: { status: 'editable', diagnostics: [] },
        children: [{
          kind: 'text',
          id: 'title-generated-id',
          name: 'Title',
          provenance: 'parsed',
          transform: { x: 1, y: 2, cx: 3, cy: 4 },
          paragraphs: [{
            align: 'center', level: 0, bullet: false,
            runs: [
              { text: 'Before', bold: true, italic: false, fontSizeHundredthPt: 3200, color: '112233', fontFamily: 'Aptos' },
              { text: ' preserved', bold: false, italic: true, fontSizeHundredthPt: 2800, color: '445566', fontFamily: 'Aptos' },
            ],
          }],
          source: { partName: 'ppt/slides/slide1.xml', objectId: 'cNvPr-2', fingerprintSha256: FINGERPRINT },
          passthrough: [],
          compatibility: { status: 'editable', diagnostics: [] },
        }],
      }, {
        kind: 'text',
        id: 'incomplete-text',
        provenance: 'parsed',
        transform: { x: 0, y: 0, cx: 1, cy: 1 },
        paragraphs: [{ runs: [{ text: 'not exact' }] }],
        source: { partName: 'ppt/slides/slide1.xml', objectId: 'cNvPr-3', fingerprintSha256: 'd'.repeat(64) },
        passthrough: [],
        compatibility: { status: 'preserveOnly', diagnostics: [] },
      }, {
        kind: 'shape',
        id: 'shape-one',
        name: 'Accent',
        provenance: 'parsed',
        transform: { x: 10, y: 20, cx: 300, cy: 200 },
        preset: 'rect',
        fill: 'DDEEFF',
        paragraphs: [],
        source: { partName: 'ppt/slides/slide1.xml', objectId: 'cNvPr-4', fingerprintSha256: '9'.repeat(64) },
        passthrough: [],
        compatibility: { status: 'editable', diagnostics: [] },
      }],
      source: { partName: 'ppt/slides/slide1.xml', objectId: 'slide-1', fingerprintSha256: 'e'.repeat(64) },
      passthrough: [],
      compatibility: { status: 'preserveOnly', diagnostics: [] },
    }],
    compatibility: { status: 'preserveOnly', diagnostics: [] },
  }
}

describe('native PPTX round trip', () => {
  it('selects exact grouped text and preserves all non-edited run formatting', () => {
    const deck = fixtureDeck()
    const targets = editablePptxTextTargets(deck)
    expect(targets).toHaveLength(1)
    expect(targets[0]).toMatchObject({ slideIndex: 0, elementName: 'Title', sourceObjectId: 'cNvPr-2' })
    expect(firstPptxRunText(targets[0])).toBe('Before')

    const mutation = buildPptxTextMutation(deck, targets[0]!, 'After', 'edit-title')
    expect(mutation).toEqual({
      expectedSourceRevision: REVISION,
      operations: [{
        operationId: 'edit-title',
        kind: 'text.replace',
        elementId: 'title-generated-id',
        expectedFingerprintSha256: FINGERPRINT,
        paragraphs: [{
          align: 'center', level: 0, bullet: false,
          runs: [
            { text: 'After', bold: true, italic: false, fontSizeHundredthPt: 3200, color: '112233', fontFamily: 'Aptos' },
            { text: ' preserved', bold: false, italic: true, fontSizeHundredthPt: 2800, color: '445566', fontFamily: 'Aptos' },
          ],
        }],
      }],
    })
  })

  it('exposes only editable exact AutoShapes and builds a complete guarded replacement', () => {
    const deck = fixtureDeck()
    const target = editablePptxShapeTargets(deck)[0]!
    expect(target).toMatchObject({ operationKind: 'autoshape.update', elementName: 'Accent', autoShape: { preset: 'rect', fill: 'DDEEFF' } })
    const mutation = buildPptxShapeMutation(deck, target, 'ellipse', '625BF6', 'edit-shape')
    expect(mutation).toEqual({
      expectedSourceRevision: REVISION,
      operations: [{
        operationId: 'edit-shape',
        kind: 'autoshape.update',
        elementId: 'shape-one',
        expectedFingerprintSha256: '9'.repeat(64),
        autoShape: { transform: target.autoShape.transform, preset: 'ellipse', fill: '625BF6' },
      }],
    })
    expect(buildPptxMutationEvidence(target, mutation)).toMatchObject({
      operation: 'autoshape.update',
      target_id: 'shape-one',
      expected_revision: REVISION,
      requested: { preset: 'ellipse', fill: '625BF6' },
    })
  })

  it('keeps an editable text target first for the primary proof flow', () => {
    expect(editablePptxTargets(fixtureDeck()).map((target) => target.operationKind)).toEqual([
      'text.replace',
      'autoshape.update',
    ])
  })

  it('does not offer preserveOnly elements even when their text is fully materialized', () => {
    const deck = fixtureDeck()
    const preserved = deck.slides[0]!.elements[1]
    if (preserved?.kind !== 'text') throw new Error('invalid fixture')
    preserved.paragraphs = [{
      align: 'left', level: 0, bullet: false,
      runs: [{ text: 'Exact but preserved', bold: false, italic: false, fontSizeHundredthPt: 2400, color: '000000', fontFamily: 'Aptos' }],
    }]
    expect(editablePptxTextTargets(deck).some((target) => target.elementId === preserved.id)).toBe(false)
  })

  it('uses the source anchor for readback even when generated ids and fingerprints change', () => {
    const before = editablePptxTextTargets(fixtureDeck())[0]!
    const afterDeck = fixtureDeck()
    const grouped = afterDeck.slides[0]!.elements[0]
    if (grouped?.kind !== 'group' || grouped.children[0]?.kind !== 'text') throw new Error('invalid fixture')
    grouped.children[0].id = 'regenerated-id'
    grouped.children[0].source!.fingerprintSha256 = 'f'.repeat(64)
    grouped.children[0].paragraphs[0]!.runs[0]!.text = 'After'

    const after = findPptxTextTarget(afterDeck, before)
    expect(after?.elementId).toBe('regenerated-id')
    expect(firstPptxRunText(after)).toBe('After')
    expect(pptxTargetKey(after!)).toBe(pptxTargetKey(before))
  })

  it('refuses no-op requests and creates safe output names', () => {
    const deck = fixtureDeck()
    const target = editablePptxTextTargets(deck)[0]!
    expect(() => buildPptxTextMutation(deck, target, 'Before', 'no-op')).toThrow(/semantic change/)
    expect(pptxDownloadName('Quarterly review.pptx')).toBe('Quarterly-review-injoffice.pptx')
    expect(pptxDownloadName('🔥.pptx')).toBe('presentation-injoffice.pptx')
  })

  it('verifies identity, revision headers, and every untouched source anchor', () => {
    const before = fixtureDeck()
    const target = editablePptxTextTargets(before)[0]!
    const after = fixtureDeck()
    const grouped = after.slides[0]!.elements[0]
    if (grouped?.kind !== 'group' || grouped.children[0]?.kind !== 'text') throw new Error('invalid fixture')
    grouped.children[0].id = 'regenerated-id'
    grouped.children[0].source!.fingerprintSha256 = 'f'.repeat(64)
    grouped.children[0].paragraphs[0]!.runs[0]!.text = 'After'
    after.sourceRevision = `rev-${'2'.repeat(64)}`

    expect(untouchedPptxElementInventory(before, target)).toHaveLength(2)
    expect(verifyPptxRoundTrip(
      before, after, target, { text: 'After' }, after.sourceRevision, `sha256:${'2'.repeat(64)}`,
      'artifact-1', 'artifact-1', false,
    )).toMatchObject({ after: 'After', artifactIdentity: 'stable', preservedElements: 2 })

    const changed = structuredClone(after)
    changed.slides[0]!.elements[2]!.source!.fingerprintSha256 = '8'.repeat(64)
    expect(() => verifyPptxRoundTrip(before, changed, target, { text: 'After' })).toThrow(/untouched element/)
  })
})
