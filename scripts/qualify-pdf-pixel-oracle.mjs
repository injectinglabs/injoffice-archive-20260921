// Original fixture, independent analytical pixel oracle. This is not an
// Acrobat/Office-export baseline and does not measure typography fidelity.
import { createHash } from 'node:crypto'
import { mkdirSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { resolve } from 'node:path'
import { PDFDocument, PDFName, PDFNumber, degrees, rgb } from 'pdf-lib'
import { createCanvas } from '@napi-rs/canvas'
import { PNG } from 'pngjs'
import { PdfViewerDocument, renderPageToCanvas } from '@injoffice/pdf/browser'
import { version as pdfjsVersion } from 'pdfjs-dist/legacy/build/pdf.mjs'

const output = process.env.SHOWCASE_OUTPUT ? resolve(process.env.SHOWCASE_OUTPUT) : mkdtempSync(resolve(tmpdir(), 'injoffice-pdf-oracle-'))
mkdirSync(output, { recursive: true })
const doc = await PDFDocument.create()
doc.setCreationDate(new Date('2026-01-01T00:00:00Z')); doc.setModificationDate(new Date('2026-01-01T00:00:00Z'))
const rectangles = [{ x: 0, y: 0, width: 40, height: 30, color: [255, 0, 0] }, { x: 40, y: 0, width: 40, height: 30, color: [0, 255, 0] }, { x: 0, y: 30, width: 40, height: 30, color: [0, 0, 255] }, { x: 40, y: 30, width: 40, height: 30, color: [255, 255, 0] }]
for (const rotation of [0, 90]) {
  const page = doc.addPage([80, 60]); page.setRotation(degrees(rotation))
  page.node.set(PDFName.of('UserUnit'), PDFNumber.of(2))
  for (const rectangle of rectangles) page.drawRectangle({ ...rectangle, color: rgb(...rectangle.color.map(value => value / 255)), borderWidth: 0 })
}
const source = await doc.save({ useObjectStreams: false })
writeFileSync(resolve(output, 'geometry.pdf'), source)
const viewer = await PdfViewerDocument.load(source)
const comparisons = []
try {
  if (viewer.pageCount !== 2) throw new Error('Expected exactly two PDF pages')
  for (let page = 1; page <= 2; page++) {
    const width = page === 1 ? 160 : 120, height = page === 1 ? 120 : 160
    const candidate = createCanvas(width, height)
    // The library only needs CSS width/height assignments from an HTML canvas.
    candidate.style = {}
    await renderPageToCanvas(viewer, page, candidate, 1, 1, { maxPixels: 20000 })
    if (candidate.width !== width || candidate.height !== height) throw new Error(`Page ${page}: rotation/UserUnit geometry drift`)
    const reference = new PNG({ width, height })
    const actual = candidate.getContext('2d').getImageData(0, 0, width, height).data
    for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
      // Analytically invert PDF's bottom-left coordinates and clockwise page
      // rotation. This does not invoke PDF.js, its viewport, or our renderer.
      const px = page === 1 ? (x + 0.5) / 2 : (y + 0.5) / 2
      const py = page === 1 ? 60 - (y + 0.5) / 2 : (x + 0.5) / 2
      const color = rectangles.find(rect => px >= rect.x && px < rect.x + rect.width && py >= rect.y && py < rect.y + rect.height).color
      const offset = (y * width + x) * 4
      reference.data.set([...color, 255], offset)
      for (let channel = 0; channel < 4; channel++) if (actual[offset + channel] !== reference.data[offset + channel]) throw new Error(`Page ${page}: pixel mismatch at (${x}, ${y}), channel ${channel}`)
    }
    const refBytes = PNG.sync.write(reference)
    writeFileSync(resolve(output, `reference-${page}.png`), refBytes)
    writeFileSync(resolve(output, `candidate-${page}.png`), candidate.toBuffer('image/png'))
    comparisons.push({ reference: `reference-${page}.png`, referenceSha256: hash(refBytes), candidate: `candidate-${page}.png` })
  }
} finally { await viewer.destroy() }
const canvasVersion = JSON.parse(readFileSync(new URL('../node_modules/@napi-rs/canvas/package.json', import.meta.url))).version
const manifest = { version: 2, cases: [{ id: 'pdf-user-unit-rotation', format: 'pdf', source: 'geometry.pdf', sourceSha256: hash(source), referenceRenderer: 'Analytical integer rectangle oracle v1 (no PDF renderer)', candidateRenderer: `InjOffice PDF viewer / PDF.js ${pdfjsVersion} / @napi-rs/canvas ${canvasVersion} / Node ${process.version}`, referenceKind: 'analytical-oracle', referenceLicense: 'Apache-2.0', referenceProvenance: 'Original generated rectangle geometry from scripts/qualify-pdf-pixel-oracle.mjs; no third-party document content', limits: { maxEncodedBytes: 1048576, maxWidth: 200, maxHeight: 200, maxPixels: 40000 }, tolerance: { perChannelDelta: 0, maxDifferentPixelsPpm: 0, maxMeanAbsoluteChannelErrorPpm: 0 }, pages: comparisons }] }
writeFileSync(resolve(output, 'manifest.json'), JSON.stringify(manifest, null, 2) + '\n')
console.log(JSON.stringify({ result: 'PASS', referenceKind: 'analytical-oracle', pages: 2, pixelsCompared: 38400, sourceSha256: hash(source), manifest: resolve(output, 'manifest.json') }, null, 2))
function hash(bytes) { return createHash('sha256').update(bytes).digest('hex') }
