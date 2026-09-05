import { PDFDocument, StandardFonts, rgb } from 'pdf-lib';
import { describe, expect, it } from 'vitest';
import { applyTextEdits } from './apply.js';
import { glyphForScalar } from './fontCmap.js';
import { applyTextInserts } from './insert.js';
import { resolveObjectMatch } from './match.js';
import { PdfiumSession } from './pdfium.js';
import { wrapText } from './rebuild.js';
import { replaceRange } from './splice.js';

describe('independently authored PDF text replacement acceptance', () => {
  it('rejects surrogate-splitting ranges and performs bounded replacement', () => {
    expect(replaceRange('A😀B', 1, 3, 'x')).toBe('AxB');
    expect(() => replaceRange('A😀B', 2, 3, 'x')).toThrow(/surrogate/i);
  });

  it('resolves complete matches and records adjacent object spans', () => {
    const extracted = {
      text: 'alpha beta',
      units: [...'alpha beta'].map((_, index) => ({ start: index, end: index + 1, object: index < 5 ? 10 : 20 })),
    };
    expect(resolveObjectMatch(extracted, 'alpha')).toEqual({
      ok: true,
      match: { start: 0, end: 5, object: 10, objects: [10], spans: [{ start: 0, end: 5, object: 10 }] },
    });
    expect(resolveObjectMatch(extracted, 'a b')).toEqual({
      ok: true,
      match: {
        start: 4,
        end: 7,
        object: 10,
        objects: [10, 20],
        spans: [{ start: 4, end: 5, object: 10 }, { start: 5, end: 7, object: 20 }],
      },
    });
  });

  it('reads a minimal standards-shaped cmap format 12', () => {
    const font = minimalFormat12Font(0x1f642, 7);
    expect(glyphForScalar(font, 0x1f642)).toBe(7);
    expect(glyphForScalar(font, 0x1f643)).toBe(0);
  });

  it('wraps deterministically at whitespace and long-token boundaries', () => {
    expect(wrapText('alpha beta', 35, 10).map((line) => line.text)).toEqual(['alpha', 'beta']);
    expect(wrapText('abcdef', 15, 10).map((line) => line.text)).toEqual(['ab', 'cd', 'ef']);
  });

  it('leaves bytes unchanged when validation fails before PDFium opens', async () => {
    const bytes = new Uint8Array([1, 2, 3]);
    const edits = await applyTextEdits(bytes, [{ page: 1, rect: [0, 0, 1, 1], oldText: '', newText: 'x' }]);
    expect(edits.applied).toBe(0);
    expect(edits.failures[0]?.code).toBe('invalid-spec');
    expect(edits.pdfBytes).toEqual(bytes);
    const inserts = await applyTextInserts(bytes, [{ page: 1, text: 'x', origin: [0, 0], fontSize: 0 }]);
    expect(inserts.applied).toBe(0);
    expect(inserts.pdfBytes).toEqual(bytes);
  });

  it('edits and inserts text through PDFium, then reloads the serialized bytes', async () => {
    const document = await PDFDocument.create();
    const page = document.addPage([300, 200]);
    const font = await document.embedFont(StandardFonts.Helvetica);
    page.drawText('alpha', { x: 40, y: 140, size: 18, font, color: rgb(0, 0, 0) });
    const original = new Uint8Array(await document.save({ useObjectStreams: false }));

    const edited = await applyTextEdits(original, [{ page: 1, rect: [0, 0, 300, 200], oldText: 'alpha', newText: 'bravo' }]);
    expect(edited.failures).toEqual([]);
    expect(edited.applied).toBe(1);
    expect(await pageText(edited.pdfBytes)).toContain('bravo');

    const inserted = await applyTextInserts(edited.pdfBytes, [{ page: 1, text: 'charlie', origin: [40, 80], fontSize: 14 }]);
    expect(inserted.failures).toEqual([]);
    expect(inserted.applied).toBe(1);
    expect(await pageText(inserted.pdfBytes)).toContain('charlie');
  });
});

async function pageText(bytes: Uint8Array): Promise<string> {
  const session = await PdfiumSession.open(bytes);
  try { return session.withTextPage(1, (_page, textPage) => session.extractText(textPage).text); }
  finally { session.close(); }
}

function minimalFormat12Font(scalar: number, glyph: number): Uint8Array {
  const bytes = new Uint8Array(60);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 0x00010000, false);
  view.setUint16(4, 1, false);
  bytes.set([0x63, 0x6d, 0x61, 0x70], 12);
  view.setUint32(20, 28, false);
  view.setUint32(24, 40, false);
  view.setUint16(28, 0, false);
  view.setUint16(30, 1, false);
  view.setUint16(32, 3, false);
  view.setUint16(34, 10, false);
  view.setUint32(36, 12, false);
  view.setUint16(40, 12, false);
  view.setUint32(44, 28, false);
  view.setUint32(52, 1, false);
  view.setUint32(56, scalar, false);
  const grown = new Uint8Array(72);
  grown.set(bytes);
  const complete = new DataView(grown.buffer);
  complete.setUint32(60, scalar, false);
  complete.setUint32(64, glyph, false);
  return grown;
}
