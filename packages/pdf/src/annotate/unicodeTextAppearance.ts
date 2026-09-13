import { appendBezierCurve, beginText, closePath, concatTransformationMatrix, defaultTextFieldAppearanceProvider, endMarkedContent, endText, fill, lineTo, moveTo, PDFHexString, PDFName, PDFOperator, PDFOperatorNames, popGraphicsState, pushGraphicsState, setTextMatrix, showText } from 'pdf-lib'
import type { PDFFont, PDFTextField, PDFWidgetAnnotation } from 'pdf-lib'
import { embeddedTextAppearance } from './embeddedTextAppearance.js'
import type { EncodedUnicodeRun, OutlineCommand } from './unicodeFontResource.js'

function outlineOperators(commands: readonly OutlineCommand[]): PDFOperator[] {
  const result: PDFOperator[] = []
  let x = 0, y = 0, startX = 0, startY = 0
  for (const { command, args: a } of commands) {
    if (command === 'moveTo') { result.push(moveTo(a[0]!, a[1]!)); x = startX = a[0]!; y = startY = a[1]! }
    else if (command === 'lineTo') { result.push(lineTo(a[0]!, a[1]!)); x = a[0]!; y = a[1]! }
    else if (command === 'quadraticCurveTo') {
      // Exact degree elevation, preserving the font's nonzero winding contours.
      result.push(appendBezierCurve(x + 2 / 3 * (a[0]! - x), y + 2 / 3 * (a[1]! - y), a[2]! + 2 / 3 * (a[0]! - a[2]!), a[3]! + 2 / 3 * (a[1]! - a[3]!), a[2]!, a[3]!))
      x = a[2]!; y = a[3]!
    } else if (command === 'bezierCurveTo') { result.push(appendBezierCurve(a[0]!, a[1]!, a[2]!, a[3]!, a[4]!, a[5]!)); x = a[4]!; y = a[5]! }
    else if (command === 'closePath') { result.push(closePath()); x = startX; y = startY }
    else throw new Error('unexpected continuation outline command')
  }
  return result
}

/** Positions every shaped glyph, retaining the provider's style and geometry. */
export function unicodeTextAppearance(field: PDFTextField, widget: PDFWidgetAnnotation, font: PDFFont, encodedRun: EncodedUnicodeRun, nominalAdvances: readonly number[]): PDFOperator[] {
  const { run, cids, encoded, outlines } = encodedRun
  if (field.getText() !== run.value && !(field.getText() === undefined && run.value === '')) throw new Error('embedded appearance has no qualified text run')
  if (field.isMultiline() || field.isCombed() || field.isPassword() || field.isFileSelector() || field.isRichFormatted()) throw new Error('embedded appearances require a plain single-line field')
  const defaultSize = (appearance: string | undefined) => {
    const matches = [...(appearance ?? '').matchAll(/\/([^\0\t\n\f\r\ ]+)[\0\t\n\f\r\ ]+(\d*\.\d+|\d+)[\0\t\n\f\r\ ]+Tf/g)]
    return matches.length ? Number(matches.at(-1)![2]) : undefined
  }
  const authoredSize = defaultSize(widget.getDefaultAppearance()) ?? defaultSize(field.acroField.getDefaultAppearance())
  const { ink } = encodedRun
  const fitInk = (authoredSize === undefined || authoredSize === 0) && (/\p{Mark}/u.test(run.value) || run.requiresActualText === true) && ink.maxY > ink.minY
  let cursor = 0
  const simple = run.glyphs.every((glyph, i) => {
    const valid = cids[i] !== 0 && glyph.x === cursor && glyph.y === 0
    cursor += glyph.advance
    return valid
  })
  if (simple && !fitInk) {
    const operators = embeddedTextAppearance(field, widget, font, { value: run.value, unitsPerEm: run.unitsPerEm, advances: run.glyphs.map(g => g.advance), nominalAdvances })
    if (!run.requiresActualText) return operators
    const replacement = PDFOperator.of(PDFOperatorNames.BeginMarkedContentSequence, [PDFName.of('Span'), field.acroField.dict.context.obj({ ActualText: PDFHexString.fromText(run.value) }).toString()])
    return operators.flatMap(operator => /(?:^|\s)(?:Tj|TJ)$/.test(operator.toString()) ? [replacement, operator, endMarkedContent()] : [operator])
  }
  const token = '\uFFFC'
  const requireToken = (value: string) => { if (value !== token) throw new Error('unexpected embedded appearance layout request') }
  const layoutFont = Object.assign(Object.create(font) as PDFFont, {
    encodeText(value: string) { requireToken(value); return encoded },
    widthOfTextAtSize(value: string, size: number) { requireToken(value); return (fitInk ? ink.maxX - ink.minX : run.width) * size / run.unitsPerEm },
    heightAtSize(size: number, options?: { descender?: boolean }) { return fitInk ? (ink.maxY - ink.minY) * size / run.unitsPerEm : font.heightAtSize(size, options) },
  })
  const layoutField = Object.assign(Object.create(field) as PDFTextField, { getText: () => token })
  const operators = defaultTextFieldAppearanceProvider(layoutField, widget, layoutFont)
  if (!Array.isArray(operators)) throw new Error('unexpected embedded appearance provider mapping')
  const textOperators = operators.filter(op => /(?:^|\s)(?:Tj|TJ|'|")$/.test(op.toString()))
  if (textOperators.length !== 1 || textOperators[0]!.toString() !== `${encoded.toString()} Tj`) throw new Error('unexpected embedded appearance text operators')
  const textIndex = operators.indexOf(textOperators[0]!)
  const prefix = operators.slice(0, textIndex).map(op => op.toString())
  const matrices = prefix.filter(op => op.endsWith(' Tm'))
  const fonts = prefix.filter(op => op.endsWith(' Tf'))
  if (matrices.length !== 1 || fonts.length !== 1) throw new Error('unexpected embedded appearance text state')
  const matrix = matrices[0]!.split(/\s+/).slice(0, -1).map(Number)
  const size = Number(fonts[0]!.split(/\s+/).at(-2))
  if (matrix.length !== 6 || !matrix.every(Number.isFinite) || !Number.isFinite(size) || size <= 0) throw new Error('invalid embedded appearance text matrix')
  const [a, b, c, d, e, f] = matrix as [number, number, number, number, number, number]
  const scale = size / run.unitsPerEm
  const replacement: PDFOperator[] = [endText(), PDFOperator.of(PDFOperatorNames.BeginMarkedContentSequence, [PDFName.of('Span'), field.acroField.dict.context.obj({ ActualText: PDFHexString.fromText(run.value) }).toString()])]
  run.glyphs.forEach((glyph, index) => {
    const x = e + (a * (glyph.x - (fitInk ? ink.minX : 0)) + c * (glyph.y - (fitInk ? ink.minY : 0))) * scale
    const y = f + (b * (glyph.x - (fitInk ? ink.minX : 0)) + d * (glyph.y - (fitInk ? ink.minY : 0))) * scale
    const cid = cids[index]!
    if (cid !== 0) replacement.push(beginText(), setTextMatrix(a, b, c, d, x, y), showText(PDFHexString.of(cid.toString(16).padStart(4, '0'))), endText())
    else replacement.push(pushGraphicsState(), concatTransformationMatrix(a * scale, b * scale, c * scale, d * scale, x, y), ...outlineOperators(outlines[index]!), fill(), popGraphicsState())
  })
  replacement.push(endMarkedContent(), beginText())
  return operators.flatMap(op => op === textOperators[0] ? replacement : [op])
}
