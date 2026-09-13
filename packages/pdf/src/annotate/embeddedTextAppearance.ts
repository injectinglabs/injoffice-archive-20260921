import { defaultTextFieldAppearanceProvider, PDFArray, PDFHexString, PDFNumber, PDFOperator, PDFOperatorNames } from 'pdf-lib'
import type { PDFFont, PDFTextField, PDFWidgetAnnotation } from 'pdf-lib'

export interface EmbeddedTextRun {
  value: string
  unitsPerEm: number
  advances: readonly number[]
  nominalAdvances: readonly number[]
}

/** Paint one qualified horizontal run without losing its positioned advances. */
export function embeddedTextAppearance(field: PDFTextField, widget: PDFWidgetAnnotation, font: PDFFont, run: EmbeddedTextRun): PDFOperator[] {
  if (field.getText() !== run.value && !(field.getText() === undefined && run.value === '')) throw new Error('embedded appearance has no qualified text run')
  if (field.isMultiline() || field.isCombed() || field.isPassword() || field.isFileSelector() || field.isRichFormatted()) {
    throw new Error('embedded appearances require a plain single-line field')
  }
  const encoded = font.encodeText(run.value)
  const hex = encoded.asString()
  if (hex.length !== run.advances.length * 4 || run.nominalAdvances.length !== run.advances.length) {
    throw new Error('embedded appearance encoding does not match its qualified run')
  }
  const width = run.advances.reduce((sum, advance) => sum + advance, 0) / run.unitsPerEm
  // The default provider's automatic sizing splits on spaces. A single local
  // layout token makes it measure the complete run, including cross-space kerns.
  // The token is never encoded or stored: the exact source encoding is supplied
  // below. Facades retain the provider's authored geometry/style/DA behavior.
  const token = '\uFFFC'
  const requireToken = (value: string) => {
    if (value !== token) throw new Error('unexpected embedded appearance layout request')
  }
  const layoutFont = Object.assign(Object.create(font) as PDFFont, {
    encodeText(value: string) { requireToken(value); return encoded },
    widthOfTextAtSize(value: string, size: number) { requireToken(value); return width * size },
  })
  const layoutField = Object.assign(Object.create(field) as PDFTextField, { getText: () => token })
  const operators = defaultTextFieldAppearanceProvider(layoutField, widget, layoutFont)
  if (!Array.isArray(operators)) throw new Error('unexpected embedded appearance provider mapping')
  const expected = `${encoded.toString()} Tj`
  const textOperators = operators.filter(operator => /(?:^|\s)(?:Tj|TJ|'|")$/.test(operator.toString()))
  if (textOperators.length !== 1 || textOperators[0]!.toString() !== expected) {
    throw new Error('unexpected embedded appearance text operators')
  }
  if (run.advances.every((advance, index) => advance === run.nominalAdvances[index])) return operators
  const adjusted = PDFArray.withContext(field.acroField.dict.context)
  for (let index = 0; index < run.advances.length; index++) {
    adjusted.push(PDFHexString.of(hex.slice(index * 4, index * 4 + 4)))
    // PDF TJ subtracts thousandths of text space from the nominal advance.
    const adjustment = (run.nominalAdvances[index]! - run.advances[index]!) * 1000 / run.unitsPerEm
    if (adjustment !== 0) adjusted.push(PDFNumber.of(adjustment))
  }
  return operators.map(operator => operator === textOperators[0]
    ? PDFOperator.of(PDFOperatorNames.ShowTextAdjusted, [adjusted]) : operator)
}
