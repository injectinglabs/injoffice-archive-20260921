import { defaultDropdownAppearanceProvider, defaultOptionListAppearanceProvider, PDFField, PDFFont, PDFCheckBox, PDFDict, PDFDocument, PDFDropdown, PDFHexString, PDFName, PDFOptionList, PDFRadioGroup, PDFTextField, StandardFonts } from 'pdf-lib'
import type { FormValueSpec } from './types.js'

export interface FormValueFailure {
  name: string
  reason: string
}

export interface FormValuesResult {
  bytes: Uint8Array
  applied: number
  skipped: FormValueFailure[]
  /** Present only when an appearance option is requested. Does not certify viewer fidelity. */
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
  /** Explicit 12pt standard-font choice artwork; every ASCII option must fit its owned widget. */
  choiceAppearance?: { font: TextAppearanceFont }
}

function qualifyOwnedAppearance(doc: PDFDocument, field: PDFField, kind: 'text' | 'choice'): void {
  // Inspect ancestors before calling helpers that recursively inherit flags.
  const ancestors = new Set<PDFDict>()
  let ancestor: PDFDict | undefined = field.acroField.dict
  while (ancestor) {
    if (ancestors.has(ancestor) || ancestors.size >= 100) throw new Error(`invalid ${kind} field ancestry`)
    ancestors.add(ancestor)
    if (ancestor.has(PDFName.of('AA')) || ancestor.has(PDFName.of('A'))) {
      throw new Error(`${kind} appearances with field actions are unsupported`)
    }
    if (!ancestor.has(PDFName.of('Parent'))) break
    const resolved: unknown = doc.context.lookup(ancestor.get(PDFName.of('Parent')))
    if (!(resolved instanceof PDFDict)) throw new Error(`invalid ${kind} field ancestry`)
    ancestor = resolved
  }
  const widgets = field.acroField.getWidgets()
  if (widgets.length === 0) throw new Error(`${kind} appearance requires a page widget`)
  const fields = doc.getForm().getFields()
  if (fields.filter(candidate => candidate.getName() === field.getName()).length !== 1) {
    throw new Error(`ambiguous ${kind} field name`)
  }
  const allWidgets = fields.flatMap(candidate => candidate.acroField.getWidgets())
  for (const widget of widgets) {
    const parent = widget.dict.get(PDFName.of('Parent'))
    const merged = widget.dict === field.acroField.dict
    if ((!merged && parent !== field.ref) || allWidgets.filter(other => other.dict === widget.dict).length !== 1) {
      throw new Error(`ambiguous ${kind} widget ownership`)
    }
    const pageOccurrences = doc.getPages().flatMap(page => page.node.Annots()?.asArray() ?? [])
      .filter(ref => doc.context.lookup(ref) === widget.dict).length
    if (pageOccurrences !== 1) throw new Error(`${kind} widget must occur on exactly one page`)
    if (widget.dict.has(PDFName.of('AA')) || widget.dict.has(PDFName.of('A'))) {
      throw new Error(`${kind} appearances with widget actions are unsupported`)
    }
    const rect = widget.getRectangle()
    if (![rect.x, rect.y, rect.width, rect.height].every(Number.isFinite) || rect.width <= 2 || rect.height <= 2) {
      throw new Error(`invalid ${kind} widget rectangle`)
    }
    const rotation = widget.getAppearanceCharacteristics()?.getRotation() ?? 0
    if (!Number.isFinite(rotation) || rotation % 90 !== 0) throw new Error(`unsupported ${kind} widget rotation`)
  }
}

function qualifyTextAppearance(doc: PDFDocument, field: PDFTextField, value: string): void {
  qualifyOwnedAppearance(doc, field, 'text')
  if (/[^\x20-\x7e]/.test(value)) throw new Error('text appearances support printable ASCII only')
  if (field.isMultiline() || field.isCombed() || field.isPassword() || field.isFileSelector() || field.isRichFormatted()) {
    throw new Error('text appearances require a plain single-line field')
  }
}

function qualifyChoiceAppearance(doc: PDFDocument, field: PDFDropdown | PDFOptionList, font: PDFFont): void {
  qualifyOwnedAppearance(doc, field, 'choice')
  if (field.isMultiselect() || (field instanceof PDFDropdown && field.isEditable())) throw new Error('choice appearances require a single-selection, noneditable field')
  const choices = field.acroField.getOptions()
  if (choices.length === 0 || choices.length > 64) throw new Error('choice appearances require 1–64 authored options')
  const exports = new Set<string>(), labels = new Set<string>()
  for (const option of choices) {
    const value = option.value.decodeText(), label = (option.display ?? option.value).decodeText()
    if (!label || label !== label.trim() || label.length > 256 || /[^\x20-\x7e]/.test(label)) throw new Error('choice appearances require printable ASCII labels up to 256 characters')
    if (exports.has(value) || labels.has(label)) throw new Error('choice appearances require unique exports and labels')
    exports.add(value); labels.add(label)
  }
  for (const widget of field.acroField.getWidgets()) {
    const authoredRotation = widget.getAppearanceCharacteristics()?.getRotation() ?? 0
    if (![0, 90, 180, 270].includes(authoredRotation)) throw new Error('choice appearances require canonical widget rotation')
    const rect = widget.getRectangle(), rotation = authoredRotation % 180
    const width = rotation === 90 ? rect.height : rect.width, height = rotation === 90 ? rect.width : rect.height
    const border = widget.getBorderStyle()?.getWidth() ?? 0
    if (!Number.isFinite(border) || border < 0) throw new Error('invalid choice widget border')
    const innerWidth = width - 2 * (border + 1), innerHeight = height - 2 * (border + 1)
    const lineHeight = font.heightAtSize(12) * 1.2
    // The list provider positions the final baseline one line-height below the
    // preceding one. Reserve descent below that baseline as well as each line.
    const descent = font.heightAtSize(12) - font.heightAtSize(12, { descender: false })
    const requiredHeight = field instanceof PDFOptionList ? choices.length * lineHeight + descent : Math.max(lineHeight, font.heightAtSize(12, { descender: false }) + 2 * descent)
    if ([...labels].some(label => font.widthOfTextAtSize(label, 12) > innerWidth) || requiredHeight > innerHeight) throw new Error('choice options do not fit the explicit 12pt appearance policy')
  }
  if (field instanceof PDFOptionList && field.acroField.dict.has(PDFName.of('TI'))) throw new Error('choice appearances with authored list scrolling are unsupported')
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
  if (!spec.value) {
    field.clear()
    return
  }
  // pdf-lib's choice helpers match display labels, not authored export values;
  // dropdown.select can also silently enable free-text editing. Preserve both
  // the source options and flags while writing the canonical export selection.
  const choices = field.acroField.getOptions()
  const matches = choices.map((option, index) => ({ option, index }))
    .filter(({ option }) => option.value.decodeText() === spec.value)
  if (matches.length > 1) throw new Error('ambiguous choice export value')
  if (matches.length === 0 && (!(field instanceof PDFDropdown) || !field.isEditable())) {
    throw new Error('choice value is not an authored export option')
  }
  const selected = matches[0]
  field.acroField.dict.set(PDFName.of('V'), selected?.option.value ?? PDFHexString.fromText(spec.value))
  if (selected) field.acroField.dict.set(PDFName.of('I'), field.acroField.dict.context.obj([selected.index]))
  else field.acroField.dict.delete(PDFName.of('I'))
}

export async function applyFormValues(bytes: Uint8Array, values: FormValueSpec[], options: FormValuesOptions = {}): Promise<FormValuesResult> {
  const appearances: FormValueAppearance[] | undefined = options.textAppearance || options.choiceAppearance ? [] : undefined
  const appearanceResult = appearances ? { appearances } : {}
  if (options.textAppearance && !['Helvetica', 'Times-Roman', 'Courier'].includes(options.textAppearance.font)) {
    throw new TypeError('unsupported text appearance font')
  }
  if (options.choiceAppearance && !['Helvetica', 'Times-Roman', 'Courier'].includes(options.choiceAppearance.font)) throw new TypeError('unsupported choice appearance font')
  if (values.length === 0) return { bytes, applied: 0, skipped: [], ...appearanceResult }
  const doc = await PDFDocument.load(bytes)
  // PDFDocument.getForm() removes XFA as a side effect. Refuse every form batch
  // before calling it, including default radio/choice edits to hybrid forms.
  if (doc.catalog.getAcroForm()?.dict.has(PDFName.of('XFA'))) {
    const reason = options.textAppearance ? 'XFA text appearances are unsupported' : 'XFA form updates are unsupported'
    return { bytes, applied: 0, skipped: values.map(spec => ({ name: spec.name, reason })), ...appearanceResult }
  }
  const form = doc.getForm()
  const skipped: FormValueFailure[] = []
  let applied = 0
  let needsViewerAppearance = false
  const priorNeedAppearances = form.acroForm.dict.get(PDFName.of('NeedAppearances'))
  const font = options.textAppearance ? doc.embedStandardFont(options.textAppearance.font as StandardFonts) : undefined

  const choiceFont = options.choiceAppearance ? doc.embedStandardFont(options.choiceAppearance.font as StandardFonts) : undefined

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
      if (choiceFont && spec.kind === 'choice') {
        if (!(field instanceof PDFDropdown || field instanceof PDFOptionList)) throw new TypeError('field is not a choice field')
        qualifyChoiceAppearance(doc, field, choiceFont)
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
    } else if (choiceFont && spec.kind === 'choice' && (field instanceof PDFDropdown || field instanceof PDFOptionList)) {
      // The explicit choice policy is fixed at 12pt; preflight proves all labels fit.
      for (const widget of field.acroField.getWidgets()) widget.setDefaultAppearance(`${widget.getDefaultAppearance() ?? field.acroField.getDefaultAppearance() ?? ''}\n/${choiceFont.name} 12 Tf`)
      const labels = field.acroField.getValues().map(value => {
        const selected = field.acroField.getOptions().find(option => option.value.decodeText() === value.decodeText())
        if (!selected) throw new Error('selected choice export has no authored label')
        return (selected.display ?? selected.value).decodeText()
      })
      // Providers paint display labels; the field dictionary keeps its exports.
      // A local facade overrides only the presentation accessor, never /V or /Opt.
      if (field instanceof PDFDropdown) field.updateAppearances(choiceFont, (source, widget, font) => defaultDropdownAppearanceProvider(Object.assign(Object.create(source) as PDFDropdown, { getSelected: () => labels }), widget, font))
      else field.updateAppearances(choiceFont, (source, widget, font) => defaultOptionListAppearanceProvider(Object.assign(Object.create(source) as PDFOptionList, { getSelected: () => labels }), widget, font))
      appearances!.push({ name: spec.name, status: 'generated', widgets: field.acroField.getWidgets().length })
    } else if (spec.kind === 'text' || spec.kind === 'choice') {
      needsViewerAppearance = true
      appearances?.push({ name: spec.name, status: 'viewer-required', widgets: field.acroField.getWidgets().length })
    }
    applied += 1
  }

  if (applied === 0) return { bytes, applied, skipped, ...appearanceResult }
  if (needsViewerAppearance) form.acroForm.dict.set(PDFName.of('NeedAppearances'), doc.context.obj(true))
  else if (priorNeedAppearances) {
    // Other, untouched fields may still require viewer regeneration.
    form.acroForm.dict.set(PDFName.of('NeedAppearances'), priorNeedAppearances)
  }
  else form.acroForm.dict.delete(PDFName.of('NeedAppearances'))
  return { bytes: await doc.save({ updateFieldAppearances: false }), applied, skipped, ...appearanceResult }
}
