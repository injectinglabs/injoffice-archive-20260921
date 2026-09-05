import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { decodePdfCollabOperation, encodePdfCollabOperation } from '../../../../packages/pdf/src/collab'
import {
  applyPdfCollabOps,
  inspectCollabPdf,
  makeCollabPdfSample,
  SAMPLE_HIGHLIGHT_QUADS,
} from './pdf'

const source = (name: string) => readFileSync(new URL(`./${name}`, import.meta.url), 'utf8')

describe('PDF annotation collab panel', () => {
  it('wires PdfPresenceManager, collab encode/decode, and annotation/form apply paths', () => {
    const panel = source('pdf.tsx')
    expect(panel).toContain('createHttpCollabTransport')
    expect(panel).toContain('new PdfPresenceManager')
    expect(panel).toContain('startSync')
    expect(panel).toContain('encodePdfCollabOperation')
    expect(panel).toContain('decodePdfCollabOperation')
    expect(panel).toContain('getPending')
    expect(panel).toContain('applyMarkups')
    expect(panel).toContain('applyNoteEdits')
    expect(panel).toContain('applyFormValues')
    expect(panel).toContain('applyAnnotDeletes')
    expect(panel).toContain("kind: 'annotation.markup'")
    expect(panel).toContain("kind: 'annotation.note'")
    expect(panel).toContain("kind: 'annotation.delete'")
    expect(panel).toContain("kind: 'form.value'")
    expect(panel).toContain('manager.publish({ page, annotLocalId:')
    expect(panel).toContain('return runExclusive')
    expect(panel).toContain('writeCollabQuery')
    expect(panel).toContain('/v1/artifacts')
    expect(panel).toContain('/v1/xlsx/extract')
    expect(panel).toMatch(/page, text, and image edits are not collaborated/i)
  })

  it('imports browser-safe annotation leaves instead of the Node-capable barrel', () => {
    const panel = source('pdf.tsx')
    expect(panel).not.toMatch(/packages\/pdf\/src\/annotate['"]/)
    expect(panel).toContain("packages/pdf/src/annotate/annotDelete'")
    expect(panel).toContain("packages/pdf/src/annotate/drawing'")
    expect(panel).toContain("packages/pdf/src/annotate/forms'")
    expect(panel).toContain("packages/pdf/src/annotate/markup'")
  })

  it('proxies opaque artifact minting through the local sidecar', () => {
    expect(source('../../vite.config.ts')).toContain("'/v1/artifacts'")
  })

  it('applies highlight, note, form, and delete ops and ignores unsupported kinds', async () => {
    const seed = await makeCollabPdfSample()
    const highlighted = await applyPdfCollabOps(seed, [
      { kind: 'page.rotate', value: { degrees: 90 } },
      encodePdfCollabOperation({
        kind: 'annotation.markup',
        value: { page: 1, type: 'highlight', color: [1, 1, 0], quads: SAMPLE_HIGHLIGHT_QUADS },
      }),
      { kind: 'form.value', value: { name: 'shared.memo', kind: 'text', value: 'hello' } },
    ])
    expect(decodePdfCollabOperation({ kind: 'page.rotate', value: {} }).ok).toBe(false)
    expect(decodePdfCollabOperation({ kind: 'content.replace', value: {} }).ok).toBe(false)
    expect(decodePdfCollabOperation({ kind: 'image.delete', value: {} }).ok).toBe(false)
    const afterMarkup = await inspectCollabPdf(highlighted)
    expect(afterMarkup.annots.filter((annot) => annot.page === 2).length).toBe(
      (await inspectCollabPdf(seed)).annots.filter((annot) => annot.page === 2).length,
    )
    expect(afterMarkup.annots.some((annot) => annot.subtype === 'highlight')).toBe(true)
    expect(afterMarkup.forms.find((field) => field.name === 'shared.memo')?.value).toBe('hello')

    const note = afterMarkup.annots.find((annot) => annot.subtype === 'note')
    expect(note).toBeTruthy()
    const edited = await applyPdfCollabOps(highlighted, [{
      kind: 'annotation.note',
      value: {
        page: note!.page,
        objNum: note!.objNum,
        rect: note!.rect,
        oldContents: note!.contents ?? '',
        contents: 'edited together',
      },
    }])
    const afterNote = await inspectCollabPdf(edited)
    expect(afterNote.annots.find((annot) => annot.subtype === 'note')?.contents).toBe('edited together')

    const highlight = afterNote.annots.find((annot) => annot.subtype === 'highlight')
    expect(highlight).toBeTruthy()
    const deleted = await applyPdfCollabOps(edited, [{
      kind: 'annotation.delete',
      value: {
        page: highlight!.page,
        objNum: highlight!.objNum,
        subtype: 'highlight',
        rect: highlight!.rect,
      },
    }])
    expect((await inspectCollabPdf(deleted)).annots.some((annot) => annot.subtype === 'highlight')).toBe(false)
  })
})
