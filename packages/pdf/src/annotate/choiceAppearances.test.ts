import { describe, it, expect } from 'vitest';
import { PDFDocument, PDFHexString, PDFName, PDFArray, PDFRawStream, PDFBool, decodePDFRawStream } from 'pdf-lib';
import { applyFormValues } from './forms.js';

async function fixture(kind: 'dropdown' | 'list' = 'dropdown') {
  const doc = await PDFDocument.create();
  const field = kind === 'dropdown' ? doc.getForm().createDropdown('choice') : doc.getForm().createOptionList('choice');
  field.acroField.setOptions(['a', 'b', 'c'].map(value => ({ value: PDFHexString.fromText(value), display: PDFHexString.fromText(value.toUpperCase()) })));
  field.addToPage(doc.addPage([300, 300]), { x: 10, y: 150, width: 100, height: 80 });
  return { doc, field };
}
const option = { choiceAppearance: { font: 'Helvetica' as const } };
const streams = (doc: PDFDocument) => doc.getForm().getField('choice').acroField.getWidgets().map(w => {
  const ap = doc.context.lookup(w.getNormalAppearance());
  if (!(ap instanceof PDFRawStream)) throw new Error('no appearance');
  return Buffer.from(decodePDFRawStream(ap).decode()).toString('latin1');
});

describe('explicit choice appearances', () => {
  it.each([-270, -90, 360, 450])('refuses noncanonical widget rotation %s before changing values', async rotation => {
    const { doc, field } = await fixture();
    field.acroField.setOptions([{ value: PDFHexString.fromText('ABCDEFGHIJK') }]);
    const widget = field.acroField.getWidgets()[0]!;
    widget.setRectangle({ x: 10, y: 10, width: 20, height: 100 });
    widget.dict.set(PDFName.of('MK'), doc.context.obj({ R: rotation }));
    const source = await doc.save({ updateFieldAppearances: false });
    const result = await applyFormValues(source, [{ name: 'choice', kind: 'choice', value: 'ABCDEFGHIJK' }], option);
    expect(result.applied).toBe(0); expect(result.bytes).toBe(source);
    expect(result.skipped[0]?.reason).toContain('canonical widget rotation');
  });

  it.each([0, 90, 180, 270])('qualifies exact quarter-turn %s using the provider dimensions', async rotation => {
    const { doc, field } = await fixture();
    field.acroField.setOptions([{ value: PDFHexString.fromText('ABCDEFGHIJK') }]);
    const widget = field.acroField.getWidgets()[0]!;
    widget.setRectangle({ x: 10, y: 10, width: rotation % 180 ? 20 : 100, height: rotation % 180 ? 100 : 20 });
    widget.dict.set(PDFName.of('MK'), doc.context.obj({ R: rotation }));
    const result = await applyFormValues(await doc.save({ updateFieldAppearances: false }), [{ name: 'choice', kind: 'choice', value: 'ABCDEFGHIJK' }], option);
    expect(result.applied).toBe(1);
    expect(streams(await PDFDocument.load(result.bytes))[0]).toContain('<4142434445464748494A4B> Tj');
  });

  it.each(['Helvetica', 'Times-Roman', 'Courier'] as const)('reserves final list descenders with %s', async fontName => {
    const { doc, field } = await fixture('list');
    field.acroField.setOptions(['g', 'q', 'y'].map(value => ({ value: PDFHexString.fromText(value) })));
    const font = doc.embedStandardFont(fontName);
    const lineHeight = font.heightAtSize(12) * 1.2;
    const descent = font.heightAtSize(12) - font.heightAtSize(12, { descender: false });
    const widget = field.acroField.getWidgets()[0]!;
    widget.getBorderStyle()!.setWidth(0);
    widget.setRectangle({ x: 10, y: 10, width: 100, height: 3 * lineHeight + 2.1 });
    const source = await doc.save({ updateFieldAppearances: false });
    const rejected = await applyFormValues(source, [{ name: 'choice', kind: 'choice', value: 'y' }], { choiceAppearance: { font: fontName } });
    expect(rejected.bytes).toBe(source); expect(rejected.applied).toBe(0);
    widget.setRectangle({ x: 10, y: 10, width: 100, height: 3 * lineHeight + descent + 2.1 });
    const accepted = await applyFormValues(await doc.save({ updateFieldAppearances: false }), [{ name: 'choice', kind: 'choice', value: 'y' }], { choiceAppearance: { font: fontName } });
    expect(accepted.applied).toBe(1);
    const stream = streams(await PDFDocument.load(accepted.bytes))[0]!;
    const baselines = [...stream.matchAll(/1 0 0 1 [\d.]+ ([\d.]+) Tm/g)].map(match => Number(match[1]));
    expect(baselines).toHaveLength(3);
    expect(Math.min(...baselines) - descent).toBeGreaterThanOrEqual(1);
  });

  for (const kind of ['dropdown', 'list'] as const) for (const font of ['Helvetica', 'Times-Roman', 'Courier'] as const) {
    it(`paints ${kind} display labels with ${font} while preserving authored exports on every widget`, async () => {
      const { doc, field } = await fixture(kind);
      field.addToPage(doc.addPage([300, 300]), { x: 20, y: 150, width: 100, height: 80 });
      const source = await doc.save(); const copy = source.slice();
      const result = await applyFormValues(source, [{ name: 'choice', kind: 'choice', value: 'b' }], { choiceAppearance: { font } });
      expect(result).toMatchObject({ applied: 1, skipped: [], appearances: [{ name: 'choice', status: 'generated', widgets: 2 }] });
      const loaded = await PDFDocument.load(result.bytes), actual = loaded.getForm().getField('choice');
      expect(actual.acroField.dict.get(PDFName.of('V'))?.toString()).toBe(PDFHexString.fromText('b').toString());
      expect(actual.acroField.getFlags()).toBe(field.acroField.getFlags());
      expect(actual.acroField.dict.get(PDFName.of('Opt'))?.toString()).toBe(field.acroField.dict.get(PDFName.of('Opt'))?.toString());
      for (const stream of streams(loaded)) {
        expect(stream).toContain('<42> Tj'); expect(stream).not.toContain('<62> Tj');
        if (kind === 'list') { expect(stream).toContain('<41> Tj'); expect(stream).toContain('<43> Tj'); expect(stream).toContain('0.6 0.7568627450980392 0.8549019607843137 rg'); }
      }
      expect(streams(loaded)).not.toEqual(streams(await PDFDocument.load(source)));
      expect(loaded.getForm().acroForm.dict.has(PDFName.of('NeedAppearances'))).toBe(false);
      expect(source).toEqual(copy);
      const cleared = await applyFormValues(result.bytes, [{ name: 'choice', kind: 'choice', value: '' }], option);
      expect(cleared.applied).toBe(1);
      expect((await PDFDocument.load(cleared.bytes)).getForm().getField('choice').acroField.dict.has(PDFName.of('V'))).toBe(false);
    });
  }

  it.each(['unicode', 'duplicate', 'many', 'long', 'multiselect', 'editable', 'scroll', 'actions', 'parent-actions', 'orphan', 'shared', 'rotation'] as const)('skips unsupported %s without changing source bytes', async kind => {
    const { doc, field } = await fixture(kind === 'scroll' ? 'list' : 'dropdown');
    if (kind === 'unicode') field.acroField.setOptions([{ value: PDFHexString.fromText('a'), display: PDFHexString.fromText('日本語') }]);
    if (kind === 'duplicate') field.acroField.setOptions(['a', 'b'].map(value => ({ value: PDFHexString.fromText(value), display: PDFHexString.fromText('Same') })));
    if (kind === 'many') field.acroField.setOptions(Array.from({ length: 65 }, (_, i) => ({ value: PDFHexString.fromText(String(i)) })));
    if (kind === 'long') field.acroField.setOptions([{ value: PDFHexString.fromText('a'), display: PDFHexString.fromText('A'.repeat(257)) }]);
    if (kind === 'multiselect') field.enableMultiselect();
    if (kind === 'editable' && 'enableEditing' in field) field.enableEditing();
    if (kind === 'scroll') field.acroField.dict.set(PDFName.of('TI'), doc.context.obj(1));
    if (kind === 'actions') field.acroField.dict.set(PDFName.of('AA'), doc.context.obj({}));
    if (kind === 'parent-actions') {
      const parent = doc.context.obj({ AA: {} }); const ref = doc.context.register(parent);
      field.acroField.dict.set(PDFName.of('Parent'), ref);
    }
    if (kind === 'orphan') doc.getPage(0).node.set(PDFName.of('Annots'), doc.context.obj([]));
    if (kind === 'shared') doc.getPage(0).node.Annots()!.push(field.acroField.dict.lookup(PDFName.of('Kids'), PDFArray).asArray()[0]!);
    if (kind === 'rotation') field.acroField.getWidgets()[0]!.dict.set(PDFName.of('MK'), doc.context.obj({ R: 45 }));
    const source = await doc.save({ updateFieldAppearances: false });
    const result = await applyFormValues(source, [{ name: 'choice', kind: 'choice', value: 'a' }], option);
    expect(result.applied).toBe(0); expect(result.bytes).toBe(source);
  });

  it('preserves prior requests and keeps default text and choice generation independent', async () => {
    const { doc } = await fixture();
    doc.getForm().createTextField('text').addToPage(doc.getPage(0), { x: 10, y: 50, width: 100, height: 20 });
    const source = await doc.save();
    const result = await applyFormValues(source, [{ name: 'choice', kind: 'choice', value: 'a' }, { name: 'text', kind: 'text', value: 'pending' }], option);
    expect(result.appearances?.map(a => a.status)).toEqual(['generated', 'viewer-required']);
    const generated = await applyFormValues(result.bytes, [{ name: 'choice', kind: 'choice', value: 'b' }], option);
    expect((await PDFDocument.load(generated.bytes)).getForm().acroForm.dict.get(PDFName.of('NeedAppearances'))).toBe(PDFBool.True);
    const noOpt = await applyFormValues(source, [{ name: 'choice', kind: 'choice', value: 'b' }]);
    expect(streams(await PDFDocument.load(noOpt.bytes))).toEqual(streams(await PDFDocument.load(source)));
  });

  it.each(['width', 'height'] as const)('refuses a list whose %s cannot fit the fixed appearance policy', async axis => {
    const { doc, field } = await fixture('list');
    field.acroField.getWidgets()[0]!.setRectangle({ x: 10, y: 10, width: axis === 'width' ? 5 : 100, height: axis === 'height' ? 10 : 80 });
    const source = await doc.save({ updateFieldAppearances: false });
    const result = await applyFormValues(source, [{ name: 'choice', kind: 'choice', value: 'a' }], option);
    expect(result.applied).toBe(0); expect(result.bytes).toBe(source); expect(result.skipped[0]?.reason).toContain('do not fit');
  });

  it('refuses XFA unchanged and rejects unsupported choice font requests', async () => {
    const { doc } = await fixture();
    doc.getForm().acroForm.dict.set(PDFName.of('XFA'), doc.context.obj('preserved'));
    const source = await doc.save({ updateFieldAppearances: false });
    const result = await applyFormValues(source, [{ name: 'choice', kind: 'choice', value: 'a' }], option);
    expect(result.bytes).toBe(source); expect(result.applied).toBe(0); expect(result.appearances).toEqual([]);
    await expect(applyFormValues(source, [], { choiceAppearance: { font: 'Bad' as 'Helvetica' } })).rejects.toThrow('unsupported choice');
  });

  it('rejects a provider failure without returning partially replaced widgets', async () => {
    const { doc, field } = await fixture();
    field.addToPage(doc.addPage(), { x: 10, y: 150, width: 100, height: 80 });
    field.acroField.getWidgets()[1]!.dict.set(PDFName.of('MK'), doc.context.obj({ BG: [2, 0, 0] }));
    const source = await doc.save({ updateFieldAppearances: false }), copy = source.slice();
    await expect(applyFormValues(source, [{ name: 'choice', kind: 'choice', value: 'a' }], option)).rejects.toThrow();
    expect(source).toEqual(copy);
  });
});
