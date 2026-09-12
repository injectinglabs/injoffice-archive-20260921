import { PDFCheckBox, PDFDict, PDFDocument, PDFDropdown, PDFName, PDFOptionList, PDFRadioGroup, PDFTextField, StandardFonts } from 'pdf-lib'
import type { FormValueSpec } from './types.js'

export interface FormValueFailure {
  name: string
  reason: string
}

export interface FormValuesResult {
  bytes: Uint8Array
  applied: number
  skipped: FormValueFailure[]
  /** Present only when textAppearance is requested. Does not certify viewer fidelity. */
  appearances?: FormValueAppearance[]
}

export interface FormValueAppearance {
  name: string
  status: 'generated' | 'viewer-required'
  widgets: number
}

export type TextAppearanceFont = 'Helvetica' | 'Times-Roman' | 'Courier'

export interface FormValuesOptions {
  /** Explicitly replaces text appearances with the selected standard font.
   * Only printable ASCII, plain single-line fields with owned page widgets qualify.
   * Unexpected appearance generation failures reject the entire operation.
   */
  textAppearance?: { font: TextAppearanceFont }
}

function qualifyTextAppearance(doc: PDFDocument, field: PDFTextField, value: string): void {
  // Inspect ancestors before calling helpers that recursively inherit flags.
  const ancestors = new Set<PDFDict>()
  let ancestor: PDFDict | undefined = field.acroField.dict
  while (ancestor) {
    if (ancestors.has(ancestor) || ancestors.size >= 100) throw new Error('invalid text field ancestry')
    ancestors.add(ancestor)
    if (ancestor.has(PDFName.of('AA')) || ancestor.has(PDFName.of('A'))) {
      throw new Error('text appearances with field actions are unsupported')
    }
    if (!ancestor.has(PDFName.of('Parent'))) break
    const resolved: unknown = doc.context.lookup(ancestor.get(PDFName.of('Parent')))
    if (!(resolved instanceof PDFDict)) throw new Error('invalid text field ancestry')
    ancestor = resolved
  }
  if (/[^\x20-\x7e]/.test(value)) throw new Error('text appearances support printable ASCII only')
  if (field.isMultiline() || field.isCombed() || field.isPassword() || field.isFileSelector() || field.isRichFormatted()) {
    throw new Error('text appearances require a plain single-line field')
  }
  if (doc.getForm().acroForm.dict.has(PDFName.of('XFA'))) throw new Error('XFA text appearances are unsupported')
  const widgets = field.acroField.getWidgets()
  if (widgets.length === 0) throw new Error('text appearance requires a page widget')
  const fields = doc.getForm().getFields()
  if (fields.filter(candidate => candidate.getName() === field.getName()).length !== 1) {
    throw new Error('ambiguous text field name')
  }
  const allWidgets = fields.flatMap(candidate => candidate.acroField.getWidgets())
  for (const widget of widgets) {
    const parent = widget.dict.get(PDFName.of('Parent'))
    const merged = widget.dict === field.acroField.dict
    if ((!merged && parent !== field.ref) || allWidgets.filter(other => other.dict === widget.dict).length !== 1) {
      throw new Error('ambiguous text widget ownership')
    }
    const pageOccurrences = doc.getPages().flatMap(page => page.node.Annots()?.asArray() ?? [])
      .filter(ref => doc.context.lookup(ref) === widget.dict).length
    if (pageOccurrences !== 1) throw new Error('text widget must occur on exactly one page')
    if (widget.dict.has(PDFName.of('AA')) || widget.dict.has(PDFName.of('A'))) {
      throw new Error('text appearances with widget actions are unsupported')
    }
    const rect = widget.getRectangle()
    if (![rect.x, rect.y, rect.width, rect.height].every(Number.isFinite) || rect.width <= 2 || rect.height <= 2) {
      throw new Error('invalid text widget rectangle')
    }
    const rotation = widget.getAppearanceCharacteristics()?.getRotation() ?? 0
    if (!Number.isFinite(rotation) || rotation % 90 !== 0) throw new Error('unsupported text widget rotation')
  }
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

export async function applyFormValues(bytes: Uint8Array, values: FormValueSpec[], options: FormValuesOptions = {}): Promise<FormValuesResult> {
  const appearances: FormValueAppearance[] | undefined = options.textAppearance ? [] : undefined
  const appearanceResult = appearances ? { appearances } : {}
  if (options.textAppearance && !['Helvetica', 'Times-Roman', 'Courier'].includes(options.textAppearance.font)) {
    throw new TypeError('unsupported text appearance font')
  }
  if (values.length === 0) return { bytes, applied: 0, skipped: [], ...appearanceResult }
  const doc = await PDFDocument.load(bytes)
  // PDFDocument.getForm() removes XFA as a side effect. Refuse the opt-in batch
  // before calling it, including mixed text/checkbox batches.
  if (options.textAppearance && doc.catalog.getAcroForm()?.dict.has(PDFName.of('XFA'))) {
    return { bytes, applied: 0, skipped: values.map(spec => ({ name: spec.name, reason: 'XFA text appearances are unsupported' })), ...appearanceResult }
  }
  const form = doc.getForm()
  const skipped: FormValueFailure[] = []
  let applied = 0
  let needsViewerAppearance = false
  const priorNeedAppearances = form.acroForm.dict.get(PDFName.of('NeedAppearances'))
  const font = options.textAppearance ? doc.embedStandardFont(options.textAppearance.font as StandardFonts) : undefined

  for (const spec of values) {
    const field = form.getFieldMaybe(spec.name)
    if (!field) {
      skipped.push({ name: spec.name, reason: 'field not found' })
      continue
    }
    try {
      if (font && spec.kind === 'text') {
        if (!(field instanceof PDFTextField)) throw new TypeError('field is not a text field')
        qualifyTextAppearance(doc, field, spec.value ?? '')
      }
      applyValue(field, spec)
    } catch (error) {
      skipped.push({ name: spec.name, reason: error instanceof Error ? error.message : String(error) })
      continue
    }
    // Do not catch this: a provider can fail after replacing one widget. Reject
    // the entire operation rather than return partially modified field data.
    if (font && field instanceof PDFTextField && spec.kind === 'text') {
      field.updateAppearances(font)
      appearances!.push({ name: spec.name, status: 'generated', widgets: field.acroField.getWidgets().length })
    } else if (spec.kind === 'text' || spec.kind === 'choice') {
      needsViewerAppearance = true
      appearances?.push({ name: spec.name, status: 'viewer-required', widgets: field.acroField.getWidgets().length })
    }
    applied += 1
  }

  if (applied === 0) return { bytes, applied, skipped, ...appearanceResult }
  if (needsViewerAppearance) form.acroForm.dict.set(PDFName.of('NeedAppearances'), doc.context.obj(true))
  else if (font && priorNeedAppearances) {
    // Other, untouched fields may still require viewer regeneration.
    form.acroForm.dict.set(PDFName.of('NeedAppearances'), priorNeedAppearances)
  }
  else form.acroForm.dict.delete(PDFName.of('NeedAppearances'))
  return { bytes: await doc.save({ updateFieldAppearances: false }), applied, skipped, ...appearanceResult }
}
