import { PDFCheckBox, PDFDocument, PDFDropdown, PDFName, PDFOptionList, PDFRadioGroup, PDFTextField } from 'pdf-lib'
import type { FormValueSpec } from './types.js'

export interface FormValueFailure {
  name: string
  reason: string
}

export interface FormValuesResult {
  bytes: Uint8Array
  applied: number
  skipped: FormValueFailure[]
}

function applyValue(field: unknown, spec: FormValueSpec): void {
  if (spec.kind === 'text') {
    if (!(field instanceof PDFTextField)) throw new TypeError('field is not a text field')
    field.setText(spec.value)
    return
  }
  if (spec.kind === 'checkbox') {
    if (!(field instanceof PDFCheckBox)) throw new TypeError('field is not a checkbox')
    if (spec.checked) field.check()
    else field.uncheck()
    return
  }
  if (spec.kind === 'radio') {
    if (!(field instanceof PDFRadioGroup)) throw new TypeError('field is not a radio group')
    if (spec.value) field.select(spec.value)
    else field.clear()
    return
  }
  if (!(field instanceof PDFDropdown || field instanceof PDFOptionList)) {
    throw new TypeError('field is not a choice field')
  }
  if (spec.value) field.select(spec.value)
  else field.clear()
}

export async function applyFormValues(bytes: Uint8Array, values: FormValueSpec[]): Promise<FormValuesResult> {
  if (values.length === 0) return { bytes, applied: 0, skipped: [] }
  const doc = await PDFDocument.load(bytes)
  const form = doc.getForm()
  const skipped: FormValueFailure[] = []
  let applied = 0
  let needsViewerAppearance = false

  for (const spec of values) {
    const field = form.getFieldMaybe(spec.name)
    if (!field) {
      skipped.push({ name: spec.name, reason: 'field not found' })
      continue
    }
    try {
      applyValue(field, spec)
      applied += 1
      needsViewerAppearance ||= spec.kind === 'text' || spec.kind === 'choice'
    } catch (error) {
      skipped.push({ name: spec.name, reason: error instanceof Error ? error.message : String(error) })
    }
  }

  if (applied === 0) return { bytes, applied, skipped }
  if (needsViewerAppearance) form.acroForm.dict.set(PDFName.of('NeedAppearances'), doc.context.obj(true))
  else form.acroForm.dict.delete(PDFName.of('NeedAppearances'))
  return { bytes: await doc.save({ updateFieldAppearances: false }), applied, skipped }
}
