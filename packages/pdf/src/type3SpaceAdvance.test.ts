import { decodePDFRawStream, PDFDict, PDFDocument, PDFName } from 'pdf-lib';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { PdfViewerDocument } from './viewer.js';
import { decideType3SpaceAdvance, ensureType3SpaceAdvances, type3SpaceCharProc } from './type3SpaceAdvance.js';

const A_GLYPH = '1000 0 0 0 700 700 d1\n0 0 700 700 re f\n';
const SPACE_GLYPH = '500 0 0 0 0 0 d1\n';

async function type3Pdf(options: {
  firstChar: number;
  lastChar: number;
  widths: number[];
  differences?: Array<number | string>;
  encodingName?: 'WinAnsiEncoding' | 'StandardEncoding';
  charProcs: Record<string, string>;
  missingWidth?: number;
  text: string;
  wordSpacing?: number;
}): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  const page = doc.addPage([300, 80]);
  const charProcs = doc.context.obj({});
  for (const [name, program] of Object.entries(options.charProcs)) {
    charProcs.set(PDFName.of(name), doc.context.register(doc.context.stream(program)));
  }
  const encoding = options.encodingName
    ? PDFName.of(options.encodingName)
    : doc.context.obj({
        Type: 'Encoding',
        Differences: (options.differences ?? []).map((item) => (typeof item === 'number' ? item : PDFName.of(item))),
      });
  const font = doc.context.obj({
    Type: 'Font',
    Subtype: 'Type3',
    FontBBox: [0, 0, 750, 750],
    FontMatrix: [0.001, 0, 0, 0.001, 0, 0],
    CharProcs: charProcs,
    Encoding: encoding,
    FirstChar: options.firstChar,
    LastChar: options.lastChar,
    Widths: options.widths,
  });
  if (options.missingWidth !== undefined) {
    font.set(PDFName.of('FontDescriptor'), doc.context.register(doc.context.obj({
      Type: 'FontDescriptor',
      FontName: 'T3',
      Flags: 4,
      FontBBox: [0, 0, 750, 750],
      ItalicAngle: 0,
      Ascent: 750,
      Descent: 0,
      MissingWidth: options.missingWidth,
    })));
  }
  const fontRef = doc.context.register(font);
  page.node.set(PDFName.of('Resources'), doc.context.obj({ Font: { F1: fontRef } }));
  const tw = options.wordSpacing ?? 0;
  page.node.set(PDFName.of('Contents'), doc.context.register(doc.context.stream(
    `BT /F1 10 Tf 10 40 Td ${tw} Tw (${options.text}) Tj ET\n`,
  )));
  return doc.save();
}

function type3Font(doc: PDFDocument): PDFDict {
  for (const [, object] of doc.context.enumerateIndirectObjects()) {
    if (object instanceof PDFDict && object.lookupMaybe(PDFName.of('Subtype'), PDFName)?.asString() === '/Type3') {
      return object;
    }
  }
  throw new Error('missing Type3 font');
}

function decoded(stream: ReturnType<typeof type3SpaceCharProc>): string {
  if (!stream) throw new Error('missing space CharProc');
  return Buffer.from(decodePDFRawStream(stream as never).decode()).toString('latin1');
}

const handles: PdfViewerDocument[] = [];
afterEach(async () => {
  await Promise.all(handles.map((doc) => doc.destroy()));
  handles.length = 0;
});

async function textWidth(bytes: Uint8Array): Promise<number> {
  const doc = await PdfViewerDocument.load(bytes);
  handles.push(doc);
  const content = await (await doc.getPage(1)).getTextContent();
  return content.items.reduce((sum, item) => sum + ('width' in item ? item.width : 0), 0);
}

describe('Type3 space advance', () => {
  it('injects a width-only /space CharProc when Widths proves character 32', async () => {
    const source = await type3Pdf({
      firstChar: 32,
      lastChar: 65,
      widths: Array.from({ length: 34 }, (_, i) => (i === 0 ? 500 : i === 33 ? 1000 : 0)),
      differences: [32, 'space', 65, 'A'],
      charProcs: { A: A_GLYPH },
      text: 'A A',
    });
    const copy = source.slice();
    expect(decideType3SpaceAdvance(type3Font(await PDFDocument.load(source)))).toEqual({
      action: 'advance', width: 500, glyphName: 'space',
    });
    const rewritten = await ensureType3SpaceAdvances(source);
    expect(source).toEqual(copy);
    expect(rewritten).not.toBe(source);
    const loaded = await PDFDocument.load(rewritten);
    const program = decoded(type3SpaceCharProc(loaded, type3Font(loaded)));
    expect(program).toBe('500 0 d0\n');
    expect(program).not.toMatch(/[fFS]|re|m |l /);
    expect(await textWidth(rewritten)).toBeGreaterThan(await textWidth(await type3Pdf({
      firstChar: 65,
      lastChar: 65,
      widths: [1000],
      differences: [65, 'A'],
      charProcs: { A: A_GLYPH },
      text: 'AA',
    })));
  });

  it('uses MissingWidth when encoding maps 32 to /space outside FirstChar', async () => {
    const source = await type3Pdf({
      firstChar: 65,
      lastChar: 65,
      widths: [1000],
      encodingName: 'WinAnsiEncoding',
      charProcs: { A: A_GLYPH },
      missingWidth: 400,
      text: 'A A',
    });
    expect(decideType3SpaceAdvance(type3Font(await PDFDocument.load(source)))).toEqual({
      action: 'advance', width: 400, glyphName: 'space',
    });
    const loaded = await PDFDocument.load(await ensureType3SpaceAdvances(source));
    expect(decoded(type3SpaceCharProc(loaded, type3Font(loaded)))).toBe('400 0 d0\n');
  });

  it('omits space when no Widths or MissingWidth proves a width', async () => {
    const source = await type3Pdf({
      firstChar: 65,
      lastChar: 65,
      widths: [1000],
      differences: [65, 'A'],
      charProcs: { A: A_GLYPH },
      text: 'A A',
    });
    expect(decideType3SpaceAdvance(type3Font(await PDFDocument.load(source)))).toEqual({
      action: 'omit', reason: 'no Widths or MissingWidth for space',
    });
    expect(await ensureType3SpaceAdvances(source)).toBe(source);
    const loaded = await PDFDocument.load(source);
    expect(type3SpaceCharProc(loaded, type3Font(loaded))).toBeUndefined();
    await expect(textWidth(source)).resolves.toBeGreaterThan(0);
  });

  it('does not replace an existing /space CharProc or invent an outline', async () => {
    const source = await type3Pdf({
      firstChar: 32,
      lastChar: 65,
      widths: Array.from({ length: 34 }, (_, i) => (i === 0 ? 500 : i === 33 ? 1000 : 0)),
      differences: [32, 'space', 65, 'A'],
      charProcs: { space: SPACE_GLYPH, A: A_GLYPH },
      text: 'A A',
    });
    expect(await ensureType3SpaceAdvances(source)).toBe(source);
    const loaded = await PDFDocument.load(source);
    expect(decoded(type3SpaceCharProc(loaded, type3Font(loaded)))).toBe(SPACE_GLYPH);
  });

  it('does not treat character 32 as space without a /space encoding', async () => {
    const source = await type3Pdf({
      firstChar: 32,
      lastChar: 32,
      widths: [500],
      differences: [32, 'gap'],
      charProcs: { A: A_GLYPH },
      text: 'A',
    });
    expect(decideType3SpaceAdvance(type3Font(await PDFDocument.load(source)))).toMatchObject({ action: 'omit' });
    expect(await ensureType3SpaceAdvances(source)).toBe(source);
  });

  it('leaves invalid bytes unchanged and does not crash the viewer', async () => {
    const garbage = new Uint8Array([1, 2, 3, 4]);
    expect(await ensureType3SpaceAdvances(garbage)).toBe(garbage);
    await expect(PdfViewerDocument.load(garbage)).rejects.toThrow();
  });
});

describe('PdfViewerDocument Type3 space', () => {
  it('loads a proven-width Type3 space without the missing-glyph warning', async () => {
    const source = await type3Pdf({
      firstChar: 32,
      lastChar: 65,
      widths: Array.from({ length: 34 }, (_, i) => (i === 0 ? 500 : i === 33 ? 1000 : 0)),
      differences: [32, 'space', 65, 'A'],
      charProcs: { A: A_GLYPH },
      text: 'A A',
    });
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    try {
      const doc = await PdfViewerDocument.load(source);
      handles.push(doc);
      await doc.getPageText(1);
      expect(warn.mock.calls.flat().join('\n')).not.toMatch(/Type3 character "space" is not available/);
    } finally {
      warn.mockRestore();
    }
  });

  it('still loads Type3 text that omits an unproven space', async () => {
    const source = await type3Pdf({
      firstChar: 65,
      lastChar: 65,
      widths: [1000],
      differences: [65, 'A'],
      charProcs: { A: A_GLYPH },
      text: 'A A',
    });
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    try {
      const doc = await PdfViewerDocument.load(source);
      handles.push(doc);
      expect(await doc.getPageText(1)).toMatch(/A/);
    } finally {
      warn.mockRestore();
    }
  });
});
