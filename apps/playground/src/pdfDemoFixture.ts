import { PDFDocument, StandardFonts, rgb, type PDFFont, type PDFPage } from 'pdf-lib'

export const PDF_DEMO_FILE_NAME = 'northstar-operating-review.pdf'
export const PDF_DEMO_PAGE_TITLES = [
  'Executive review',
  'Operating metrics',
  'Approval memo',
  'Review and markup',
] as const

const INK = rgb(0.09, 0.12, 0.2)
const MUTED = rgb(0.35, 0.4, 0.49)
const INDIGO = rgb(0.25, 0.22, 0.75)
const CORAL = rgb(0.95, 0.38, 0.31)
const PALE = rgb(0.95, 0.96, 1)
const LINE = rgb(0.84, 0.86, 0.91)

function drawChrome(page: PDFPage, title: string, pageNumber: number, regular: PDFFont, bold: PDFFont) {
  page.drawRectangle({ x: 0, y: 720, width: 612, height: 72, color: INK })
  page.drawText('NORTHSTAR / OPERATING REVIEW', { x: 48, y: 760, size: 8, font: bold, color: rgb(0.72, 0.75, 0.86) })
  page.drawText(title, { x: 48, y: 734, size: 19, font: bold, color: rgb(1, 1, 1) })
  page.drawLine({ start: { x: 48, y: 48 }, end: { x: 564, y: 48 }, thickness: 0.8, color: LINE })
  page.drawText('Confidential · September 2026', { x: 48, y: 30, size: 8, font: regular, color: MUTED })
  page.drawText(`${pageNumber} / ${PDF_DEMO_PAGE_TITLES.length}`, { x: 534, y: 30, size: 8, font: bold, color: MUTED })
}

function drawKpi(page: PDFPage, x: number, label: string, value: string, detail: string, regular: PDFFont, bold: PDFFont) {
  page.drawRectangle({ x, y: 568, width: 156, height: 96, color: PALE, borderColor: LINE, borderWidth: 0.7 })
  page.drawText(label.toUpperCase(), { x: x + 14, y: 640, size: 8, font: bold, color: MUTED })
  page.drawText(value, { x: x + 14, y: 606, size: 25, font: bold, color: INDIGO })
  page.drawText(detail, { x: x + 14, y: 585, size: 8.5, font: regular, color: MUTED })
}

function drawBullet(page: PDFPage, y: number, title: string, detail: string, regular: PDFFont, bold: PDFFont) {
  page.drawCircle({ x: 56, y: y + 4, size: 3.2, color: CORAL })
  page.drawText(title, { x: 70, y, size: 11, font: bold, color: INK })
  page.drawText(detail, { x: 70, y: y - 17, size: 9, font: regular, color: MUTED })
}

export async function createPdfDemoFixture(): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  doc.setTitle('Northstar operating review')
  doc.setAuthor('InjOffice demo')
  doc.setSubject('A deterministic, browser-generated PDF workbench fixture')
  doc.setKeywords(['InjOffice', 'operating review', 'approval'])
  doc.setCreator('InjOffice playground')
  doc.setProducer('InjOffice playground with pdf-lib')
  const fixtureDate = new Date('2026-09-01T12:00:00.000Z')
  doc.setCreationDate(fixtureDate)
  doc.setModificationDate(fixtureDate)

  const regular = await doc.embedFont(StandardFonts.Helvetica)
  const bold = await doc.embedFont(StandardFonts.HelveticaBold)

  const executive = doc.addPage([612, 792])
  drawChrome(executive, PDF_DEMO_PAGE_TITLES[0], 1, regular, bold)
  executive.drawText('Launch health', { x: 48, y: 690, size: 10, font: bold, color: CORAL })
  executive.drawText('Momentum is strong; reliability stays the gate.', { x: 48, y: 674, size: 13, font: regular, color: INK })
  drawKpi(executive, 48, 'Activation', '128%', 'vs. quarterly target', regular, bold)
  drawKpi(executive, 228, 'Time-to-value', '7 days', 'down from 11 days', regular, bold)
  drawKpi(executive, 408, 'Pilot conversion', '64%', 'enterprise cohort', regular, bold)
  executive.drawText('Executive priorities', { x: 48, y: 526, size: 14, font: bold, color: INK })
  drawBullet(executive, 492, 'Protect the onboarding gain', 'Keep adaptive setup behind the same reliability budget.', regular, bold)
  drawBullet(executive, 438, 'Close trust gaps before scale', 'Audit-log export and accessibility findings remain release gates.', regular, bold)
  drawBullet(executive, 384, 'Make ownership visible', 'Every pilot has a named product, success, and engineering owner.', regular, bold)
  executive.drawRectangle({ x: 48, y: 246, width: 516, height: 80, color: rgb(1, 0.96, 0.94) })
  executive.drawText('RECOMMENDATION', { x: 64, y: 300, size: 8, font: bold, color: CORAL })
  executive.drawText('Approve the October rollout with weekly adoption', { x: 64, y: 277, size: 13, font: bold, color: INK })
  executive.drawText('and reliability checkpoints.', { x: 64, y: 258, size: 13, font: bold, color: INK })

  const metrics = doc.addPage([612, 792])
  drawChrome(metrics, PDF_DEMO_PAGE_TITLES[1], 2, regular, bold)
  metrics.drawText('Four-week operating scorecard', { x: 48, y: 684, size: 14, font: bold, color: INK })
  const columns = [48, 234, 326, 416, 506]
  const headers = ['Cohort', 'Activated', '7-day use', 'Conversion', 'Tickets']
  metrics.drawRectangle({ x: 48, y: 638, width: 516, height: 28, color: INK })
  headers.forEach((header, index) => metrics.drawText(header, { x: columns[index]!, y: 648, size: 8.5, font: bold, color: rgb(1, 1, 1) }))
  const rows = [
    ['Enterprise', '82%', '71%', '64%', '18'],
    ['Mid-market', '76%', '66%', '58%', '24'],
    ['Self-serve', '69%', '61%', '—', '31'],
  ]
  rows.forEach((row, rowIndex) => {
    const y = 610 - rowIndex * 38
    if (rowIndex % 2 === 0) metrics.drawRectangle({ x: 48, y: y - 9, width: 516, height: 30, color: PALE })
    row.forEach((value, columnIndex) => metrics.drawText(value, { x: columns[columnIndex]!, y, size: 9.5, font: columnIndex === 0 ? bold : regular, color: INK }))
  })
  metrics.drawText('Reliability budget', { x: 48, y: 454, size: 14, font: bold, color: INK })
  const bars = [
    ['API availability', 0.997, '99.7%'],
    ['Successful imports', 0.984, '98.4%'],
    ['Onboarding completion', 0.82, '82.0%'],
  ] as const
  bars.forEach(([label, ratio, value], index) => {
    const y = 408 - index * 62
    metrics.drawText(label, { x: 48, y: y + 18, size: 9, font: bold, color: INK })
    metrics.drawText(value, { x: 520, y: y + 18, size: 9, font: bold, color: INDIGO })
    metrics.drawRectangle({ x: 48, y, width: 516, height: 9, color: LINE })
    metrics.drawRectangle({ x: 48, y, width: 516 * ratio, height: 9, color: INDIGO })
  })
  metrics.drawText('Source: repository-owned synthetic demo data. No customer data.', { x: 48, y: 174, size: 8.5, font: regular, color: MUTED })

  const approval = doc.addPage([612, 792])
  drawChrome(approval, PDF_DEMO_PAGE_TITLES[2], 3, regular, bold)
  approval.drawText('Release decision', { x: 48, y: 682, size: 14, font: bold, color: INK })
  approval.drawText('Complete the fields, apply values, then download the edited PDF.', { x: 48, y: 661, size: 10, font: regular, color: MUTED })
  const form = doc.getForm()
  approval.drawText('Decision owner', { x: 48, y: 614, size: 9, font: bold, color: INK })
  form.createTextField('release.decisionOwner').addToPage(approval, { x: 48, y: 578, width: 250, height: 26, borderColor: LINE, backgroundColor: rgb(1, 1, 1) })
  approval.drawText('Rollout tier', { x: 326, y: 614, size: 9, font: bold, color: INK })
  const tier = form.createDropdown('release.rolloutTier')
  tier.addOptions(['Pilot only', '25% staged', 'General availability'])
  tier.select('25% staged')
  tier.addToPage(approval, { x: 326, y: 578, width: 238, height: 26, borderColor: LINE, backgroundColor: rgb(1, 1, 1) })
  approval.drawText('Decision notes', { x: 48, y: 536, size: 9, font: bold, color: INK })
  const notes = form.createTextField('release.decisionNotes')
  notes.enableMultiline()
  notes.addToPage(approval, { x: 48, y: 386, width: 516, height: 136, borderColor: LINE, backgroundColor: PALE })
  form.createCheckBox('release.approved').addToPage(approval, { x: 48, y: 338, width: 18, height: 18, borderColor: INDIGO, backgroundColor: rgb(1, 1, 1) })
  approval.drawText('I approve the staged rollout and its weekly checkpoints.', { x: 78, y: 342, size: 10, font: bold, color: INK })
  approval.drawText('Suggested evidence', { x: 48, y: 280, size: 10, font: bold, color: CORAL })
  approval.drawText('• audit-log export release note', { x: 48, y: 254, size: 9.5, font: regular, color: INK })
  approval.drawText('• accessibility closeout report', { x: 48, y: 234, size: 9.5, font: regular, color: INK })
  approval.drawText('• pilot owner and rollback roster', { x: 48, y: 214, size: 9.5, font: regular, color: INK })

  const review = doc.addPage([612, 792])
  drawChrome(review, PDF_DEMO_PAGE_TITLES[3], 4, regular, bold)
  review.drawText('Markup exercise', { x: 48, y: 682, size: 14, font: bold, color: INK })
  review.drawText('Search for “weekly checkpoints,” then highlight, underline,', { x: 48, y: 653, size: 11, font: regular, color: INK })
  review.drawText('or strike the sentence below. Add a note and an APPROVED stamp.', { x: 48, y: 635, size: 11, font: regular, color: INK })
  review.drawRectangle({ x: 48, y: 520, width: 516, height: 76, color: PALE, borderColor: INDIGO, borderWidth: 1 })
  review.drawText('The October rollout may proceed with weekly checkpoints', { x: 66, y: 564, size: 12, font: bold, color: INK })
  review.drawText('for adoption, reliability, accessibility, and rollback readiness.', { x: 66, y: 543, size: 11, font: regular, color: INK })
  review.drawText('Reviewer notes', { x: 48, y: 468, size: 10, font: bold, color: MUTED })
  ;[432, 394, 356, 318].forEach((y) => review.drawLine({ start: { x: 48, y }, end: { x: 564, y }, thickness: 0.6, color: LINE }))
  review.drawRectangle({ x: 386, y: 198, width: 178, height: 72, borderColor: CORAL, borderWidth: 1.2 })
  review.drawText('STAMP AREA', { x: 438, y: 238, size: 9, font: bold, color: CORAL })
  review.drawText('Apply APPROVED here', { x: 416, y: 218, size: 9, font: regular, color: MUTED })

  return doc.save({ useObjectStreams: false })
}
