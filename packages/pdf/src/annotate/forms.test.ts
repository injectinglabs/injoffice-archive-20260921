import { decodePDFRawStream, PDFArray, PDFBool, PDFDict, PDFDocument, PDFHexString, PDFName, PDFRawStream } from 'pdf-lib';
import { describe, expect, it } from 'vitest';
import { applyFormValues } from './forms.js';
import type { FormValueSpec } from './types.js';

async function formDoc(): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  const page = doc.addPage([300, 300]);
  const form = doc.getForm();
  form.createTextField('name').addToPage(page, { x: 10, y: 200, width: 100, height: 20 });
  form.createCheckBox('agree').addToPage(page, { x: 10, y: 150, width: 20, height: 20 });
  const rg = form.createRadioGroup('color');
  rg.addOptionToPage('red', page, { x: 10, y: 100, width: 20, height: 20 });
  rg.addOptionToPage('blue', page, { x: 40, y: 100, width: 20, height: 20 });
  const dd = form.createDropdown('country');
  dd.addOptions(['US', 'UK']);
  dd.addToPage(page, { x: 10, y: 50, width: 100, height: 20 });
  return doc.save();
}

describe('applyFormValues', () => {
  it('refuses a default mixed XFA form update before changing any source bytes', async () => {
    const doc = await PDFDocument.load(await formDoc());
    doc.getForm().acroForm.dict.set(PDFName.of('XFA'), doc.context.obj('preserve source XFA'));
    const source = await doc.save({ updateFieldAppearances: false });
    const copy = source.slice();
    const result = await applyFormValues(source, [
      { name: 'color', kind: 'radio', value: 'blue' },
      { name: 'country', kind: 'choice', value: 'UK' },
      { name: 'name', kind: 'text', value: 'AFTER' },
      { name: 'agree', kind: 'checkbox', checked: true },
    ]);
    expect(result.applied).toBe(0);
    expect(result.skipped).toHaveLength(4);
    expect(result.skipped.every(item => item.reason === 'XFA form updates are unsupported')).toBe(true);
    expect(result.bytes).toBe(source);
    expect(result.bytes).toEqual(copy);
    expect((await PDFDocument.load(result.bytes)).catalog.getAcroForm()!.dict.has(PDFName.of('XFA'))).toBe(true);
  });

  const portable = { textAppearance: { font: 'Helvetica' as const } };
  const widgetStreams = (doc: PDFDocument, name: string) => doc.getForm().getTextField(name).acroField.getWidgets()
    .map(widget => {
      const stream = doc.context.lookup(widget.getNormalAppearance());
      if (!(stream instanceof PDFRawStream)) throw new Error('expected raw appearance stream');
      return Buffer.from(decodePDFRawStream(stream).decode()).toString('latin1');
    });

  it('regenerates every owned widget, preserves unrelated appearances and source bytes', async () => {
    const doc = await PDFDocument.load(await formDoc());
    const field = doc.getForm().getTextField('name');
    field.setText('BEFORE');
    field.addToPage(doc.addPage([300, 300]), { x: 10, y: 200, width: 100, height: 20 });
    doc.getForm().createTextField('other').addToPage(doc.getPage(0), { x: 10, y: 250, width: 100, height: 20 });
    const source = await doc.save();
    const sourceCopy = source.slice();
    const original = await PDFDocument.load(source);
    const result = await applyFormValues(source, [{ name: 'name', kind: 'text', value: 'AFTER' }], portable);
    expect(result).toMatchObject({ applied: 1, skipped: [], appearances: [{ name: 'name', status: 'generated', widgets: 2 }] });
    const loaded = await PDFDocument.load(result.bytes);
    expect(loaded.getForm().getTextField('name').getText()).toBe('AFTER');
    for (const stream of widgetStreams(loaded, 'name')) expect(stream).toContain('<4146544552> Tj');
    expect(widgetStreams(loaded, 'name')).not.toEqual(widgetStreams(original, 'name'));
    expect(widgetStreams(loaded, 'other')).toEqual(widgetStreams(original, 'other'));
    expect(loaded.getForm().getTextField('other').getText()).toBeUndefined();
    expect(loaded.getForm().acroForm.dict.has(PDFName.of('NeedAppearances'))).toBe(false);
    expect(source).toEqual(sourceCopy);
  });

  it.each(['日本語', 'مرحبا', 'café', 'line\nbreak', '\t', '\u007f'])('skips unsupported %s without changing its value/AP in a mixed batch', async value => {
    const source = await formDoc();
    const result = await applyFormValues(source, [
      { name: 'name', kind: 'text', value }, { name: 'agree', kind: 'checkbox', checked: true },
    ], portable);
    expect(result).toMatchObject({ applied: 1, skipped: [{ name: 'name', reason: 'text appearances support printable ASCII only' }], appearances: [] });
    const loaded = await PDFDocument.load(result.bytes);
    expect(loaded.getForm().getTextField('name').getText()).toBeUndefined();
    expect(widgetStreams(loaded, 'name')).toEqual(widgetStreams(await PDFDocument.load(source), 'name'));
  });

  it.each(['enableMultiline', 'enablePassword', 'enableFileSelection', 'enableRichFormatting', 'enableCombing'] as const)(
    'skips unsupported field flag %s', async flag => {
      const doc = await PDFDocument.load(await formDoc());
      const field = doc.getForm().getTextField('name');
      field.setMaxLength(8);
      field[flag]();
      const source = await doc.save({ updateFieldAppearances: false });
      const result = await applyFormValues(source, [{ name: 'name', kind: 'text', value: 'AFTER' }], portable);
      expect(result.applied).toBe(0);
      expect(result.bytes).toBe(source);
      expect(result.skipped[0]?.reason).toContain('plain single-line');
    },
  );

  it.each(['Helvetica', 'Times-Roman', 'Courier'] as const)('supports explicit %s and empty values', async font => {
    const result = await applyFormValues(await formDoc(), [{ name: 'name', kind: 'text', value: '' }], { textAppearance: { font } });
    expect(result.applied).toBe(1);
    const loaded = await PDFDocument.load(result.bytes);
    expect(loaded.getForm().getTextField('name').getText() ?? '').toBe('');
    expect(widgetStreams(loaded, 'name')[0]).toContain('Tf');
  });

  it('retains a prior NeedAppearances request and reports choice appearances separately', async () => {
    const doc = await PDFDocument.load(await formDoc());
    doc.getForm().acroForm.dict.set(PDFName.of('NeedAppearances'), PDFBool.True);
    const source = await doc.save();
    const text = { name: 'name', kind: 'text' as const, value: 'AFTER' };
    const result = await applyFormValues(source, [text], portable);
    expect((await PDFDocument.load(result.bytes)).getForm().acroForm.dict.get(PDFName.of('NeedAppearances'))).toBe(PDFBool.True);
    const checkboxOnly = await applyFormValues(source, [{ name: 'agree', kind: 'checkbox', checked: true }], portable);
    expect((await PDFDocument.load(checkboxOnly.bytes)).getForm().acroForm.dict.get(PDFName.of('NeedAppearances'))).toBe(PDFBool.True);
    const mixed = await applyFormValues(await formDoc(), [text, { name: 'country', kind: 'choice', value: 'UK' }], portable);
    expect(mixed.appearances?.map(item => item.status)).toEqual(['generated', 'viewer-required']);
  });

  it.each([
    { name: 'agree', kind: 'checkbox' as const, checked: true },
    { name: 'color', kind: 'radio' as const, value: 'blue' },
  ])('retains pending text and choice appearances after a separate $kind edit', async spec => {
    const pending = await applyFormValues(await formDoc(), [
      { name: 'name', kind: 'text', value: '日本語' },
      { name: 'country', kind: 'choice', value: 'UK' },
    ]);
    const sourceCopy = pending.bytes.slice();
    const before = await PDFDocument.load(pending.bytes);
    const result = await applyFormValues(pending.bytes, [spec]);
    expect(result).toMatchObject({ applied: 1, skipped: [] });
    const loaded = await PDFDocument.load(result.bytes);
    const form = loaded.getForm();
    expect(form.acroForm.dict.get(PDFName.of('NeedAppearances'))).toBe(PDFBool.True);
    expect(form.getTextField('name').getText()).toBe('日本語');
    expect(form.getDropdown('country').getSelected()).toEqual(['UK']);
    expect(widgetStreams(loaded, 'name')).toEqual(widgetStreams(before, 'name'));
    if (spec.kind === 'checkbox') expect(form.getCheckBox('agree').isChecked()).toBe(true);
    else expect(form.getRadioGroup('color').getSelected()).toBe('blue');
    expect(pending.bytes).toEqual(sourceCopy);
  });

  it.each(['XFA', 'actions', 'shared', 'orphan', 'rotation'] as const)('skips %s forms/widgets', async kind => {
    const doc = await PDFDocument.load(await formDoc());
    const form = doc.getForm();
    const field = form.getTextField('name');
    if (kind === 'XFA') form.acroForm.dict.set(PDFName.of('XFA'), doc.context.obj('unsupported'));
    if (kind === 'actions') field.acroField.dict.set(PDFName.of('AA'), doc.context.obj({}));
    if (kind === 'shared') doc.getPage(0).node.Annots()!.push(field.acroField.dict.lookup(PDFName.of('Kids'), PDFArray).asArray()[0]!);
    if (kind === 'orphan') doc.getPage(0).node.set(PDFName.of('Annots'), doc.context.obj([]));
    if (kind === 'rotation') field.acroField.getWidgets()[0]!.dict.set(PDFName.of('MK'), doc.context.obj({ R: 45 }));
    const source = await doc.save({ updateFieldAppearances: false });
    const result = await applyFormValues(source, [{ name: 'name', kind: 'text', value: 'AFTER' }], portable);
    expect(result.applied).toBe(0);
    expect(result.bytes).toBe(source);
  });

  it('rejects the whole operation on an unexpected provider failure after another field succeeds', async () => {
    const doc = await PDFDocument.load(await formDoc());
    const field = doc.getForm().getTextField('name');
    field.addToPage(doc.addPage([300, 300]), { x: 10, y: 200, width: 100, height: 20 });
    // An out-of-range background color is read by the provider after setText.
    // The first widget has already regenerated when the second widget fails.
    field.acroField.getWidgets()[1]!.dict.set(PDFName.of('MK'), doc.context.obj({ BG: [2, 0, 0] }));
    const source = await doc.save({ updateFieldAppearances: false });
    const copy = source.slice();
    await expect(applyFormValues(source, [
      { name: 'agree', kind: 'checkbox', checked: true }, { name: 'name', kind: 'text', value: 'AFTER' },
    ], portable)).rejects.toThrow();
    expect(source).toEqual(copy);
  });

  it('retains text and AP when max length rejects a mixed opt-in request', async () => {
    const doc = await PDFDocument.load(await formDoc());
    doc.getForm().getTextField('name').setMaxLength(2);
    const source = await doc.save();
    const result = await applyFormValues(source, [
      { name: 'name', kind: 'text', value: 'LONG' }, { name: 'agree', kind: 'checkbox', checked: true },
    ], portable);
    expect(result.applied).toBe(1);
    const loaded = await PDFDocument.load(result.bytes);
    expect(loaded.getForm().getTextField('name').getText()).toBeUndefined();
    expect(widgetStreams(loaded, 'name')).toEqual(widgetStreams(await PDFDocument.load(source), 'name'));
  });

  it('refuses actions on a non-terminal field ancestor', async () => {
    const doc = await PDFDocument.create();
    const field = doc.getForm().createTextField('group.name');
    field.addToPage(doc.addPage(), { x: 10, y: 10, width: 100, height: 20 });
    field.acroField.dict.lookup(PDFName.of('Parent'), PDFDict).set(PDFName.of('AA'), doc.context.obj({}));
    const source = await doc.save();
    const result = await applyFormValues(source, [{ name: 'group.name', kind: 'text', value: 'AFTER' }], portable);
    expect(result.bytes).toBe(source);
    expect(result.skipped).toEqual([{ name: 'group.name', reason: 'text appearances with field actions are unsupported' }]);
  });

  it('fills a text field', async () => {
    const bytes = await formDoc();
    const result = await applyFormValues(bytes, [{ name: 'name', kind: 'text', value: 'Nick' }]);
    expect(result.skipped).toEqual([]);
    expect(result.applied).toBe(1);
    const loaded = await PDFDocument.load(result.bytes);
    expect(loaded.getForm().getTextField('name').getText()).toBe('Nick');
    expect(loaded.getForm().acroForm.dict.lookup(PDFName.of('NeedAppearances'), PDFBool).asBoolean()).toBe(true);
  });

  it.each(['\u65e5\u672c\u8a9e', '\u0645\u0631\u062d\u0628\u0627', '\u0928\u092e\u0938\u094d\u0924\u0947'])(
    'preserves %s text and requests viewer appearances',
    async (value) => {
      const result = await applyFormValues(await formDoc(), [{ name: 'name', kind: 'text', value }]);
      expect(result).toMatchObject({ applied: 1, skipped: [] });
      const loaded = await PDFDocument.load(result.bytes);
      expect(loaded.getForm().getTextField('name').getText()).toBe(value);
      expect(loaded.getForm().acroForm.dict.lookup(PDFName.of('NeedAppearances'), PDFBool).asBoolean()).toBe(true);
    },
  );

  it('checks and unchecks a checkbox', async () => {
    const bytes = await formDoc();
    const checked = await applyFormValues(bytes, [{ name: 'agree', kind: 'checkbox', checked: true }]);
    expect((await PDFDocument.load(checked.bytes)).getForm().getCheckBox('agree').isChecked()).toBe(true);
    const unchecked = await applyFormValues(checked.bytes, [{ name: 'agree', kind: 'checkbox', checked: false }]);
    expect((await PDFDocument.load(unchecked.bytes)).getForm().getCheckBox('agree').isChecked()).toBe(false);
    expect((await PDFDocument.load(unchecked.bytes)).getForm().acroForm.dict.get(PDFName.of('NeedAppearances'))).toBeUndefined();
  });

  it('selects a radio group option, and clears it with an empty value', async () => {
    const bytes = await formDoc();
    const selected = await applyFormValues(bytes, [{ name: 'color', kind: 'radio', value: 'blue' }]);
    expect((await PDFDocument.load(selected.bytes)).getForm().getRadioGroup('color').getSelected()).toBe('blue');
    const cleared = await applyFormValues(selected.bytes, [{ name: 'color', kind: 'radio' }]);
    expect((await PDFDocument.load(cleared.bytes)).getForm().getRadioGroup('color').getSelected()).toBeUndefined();
  });

  it('selects a choice (dropdown) option', async () => {
    const bytes = await formDoc();
    const result = await applyFormValues(bytes, [{ name: 'country', kind: 'choice', value: 'UK' }]);
    expect(result.applied).toBe(1);
    expect((await PDFDocument.load(result.bytes)).getForm().getDropdown('country').getSelected()).toEqual(['UK']);
  });

  it.each(['dropdown', 'list'] as const)('selects authored %s export values without changing field flags or options', async kind => {
    const doc = await PDFDocument.create();
    const form = doc.getForm();
    const field = kind === 'dropdown' ? form.createDropdown('choice') : form.createOptionList('choice');
    field.acroField.setOptions([{ value: PDFHexString.fromText('e'), display: PDFHexString.fromText('E') }]);
    field.addToPage(doc.addPage(), { x: 10, y: 10, width: 100, height: 30 });
    const source = await doc.save();
    const original = await PDFDocument.load(source);
    const originalField = original.getForm().getField('choice');
    const result = await applyFormValues(source, [{ name: 'choice', kind: 'choice', value: 'e' }]);
    expect(result).toMatchObject({ applied: 1, skipped: [] });
    const loaded = await PDFDocument.load(result.bytes);
    const actual = loaded.getForm().getField('choice');
    expect(actual.acroField.dict.get(PDFName.of('V'))?.toString()).toBe(PDFHexString.fromText('e').toString());
    expect(actual.acroField.getFlags()).toBe(originalField.acroField.getFlags());
    expect(actual.acroField.dict.get(PDFName.of('Opt'))?.toString()).toBe(originalField.acroField.dict.get(PDFName.of('Opt'))?.toString());
    expect(actual.acroField.dict.lookup(PDFName.of('I'), PDFArray).asArray().map(v => v.toString())).toEqual(['0']);
    expect(loaded.getForm().acroForm.dict.get(PDFName.of('NeedAppearances'))).toBe(PDFBool.True);
    const invalid = await applyFormValues(source, [{ name: 'choice', kind: 'choice', value: 'E' }]);
    expect(invalid.applied).toBe(0); expect(invalid.bytes).toBe(source);
  });

  it('allows free text only in an already editable dropdown and rejects ambiguous exports', async () => {
    const doc = await PDFDocument.load(await formDoc());
    const field = doc.getForm().getDropdown('country');
    const source = await doc.save();
    const rejected = await applyFormValues(source, [{ name: 'country', kind: 'choice', value: 'OTHER' }]);
    expect(rejected.applied).toBe(0); expect(rejected.bytes).toBe(source);
    field.enableEditing();
    const editable = await applyFormValues(await doc.save(), [{ name: 'country', kind: 'choice', value: 'OTHER' }]);
    expect((await PDFDocument.load(editable.bytes)).getForm().getDropdown('country').getSelected()).toEqual(['OTHER']);
    field.acroField.setOptions([{ value: PDFHexString.fromText('same') }, { value: PDFHexString.fromText('same') }]);
    const ambiguous = await doc.save({ updateFieldAppearances: false });
    const result = await applyFormValues(ambiguous, [{ name: 'country', kind: 'choice', value: 'same' }]);
    expect(result.applied).toBe(0); expect(result.bytes).toBe(ambiguous);
  });

  it('sets multiple fields in one batch', async () => {
    const bytes = await formDoc();
    const values: FormValueSpec[] = [
      { name: 'name', kind: 'text', value: 'Nick' },
      { name: 'agree', kind: 'checkbox', checked: true },
      { name: 'color', kind: 'radio', value: 'red' },
    ];
    const result = await applyFormValues(bytes, values);
    expect(result.applied).toBe(3);
    expect(result.skipped).toEqual([]);
  });

  it('reports a missing field as skipped, not a thrown error, and still applies the others', async () => {
    const bytes = await formDoc();
    const result = await applyFormValues(bytes, [
      { name: 'name', kind: 'text', value: 'Nick' },
      { name: 'does-not-exist', kind: 'text', value: 'x' },
    ]);
    expect(result.applied).toBe(1);
    expect(result.skipped).toHaveLength(1);
    expect(result.skipped[0]?.name).toBe('does-not-exist');
    expect((await PDFDocument.load(result.bytes)).getForm().getTextField('name').getText()).toBe('Nick');
  });

  it('reports a wrong-kind field as skipped (a text field name used as choice)', async () => {
    const bytes = await formDoc();
    const result = await applyFormValues(bytes, [{ name: 'name', kind: 'choice', value: 'x' }]);
    expect(result.applied).toBe(0);
    expect(result.skipped).toHaveLength(1);
  });

  it('is a no-op that returns the same bytes for an empty list', async () => {
    const bytes = await formDoc();
    const result = await applyFormValues(bytes, []);
    expect(result.bytes).toBe(bytes);
    expect(result.applied).toBe(0);
  });

  it('returns the original bytes untouched when every field fails', async () => {
    const bytes = await formDoc();
    const result = await applyFormValues(bytes, [{ name: 'nope', kind: 'text', value: 'x' }]);
    expect(result.bytes).toBe(bytes);
    expect(result.applied).toBe(0);
  });
});
