import { PDFBool, PDFDocument, PDFName } from 'pdf-lib';
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
