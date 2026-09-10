import assert from 'node:assert/strict'
import { test } from 'node:test'
import { PDFDocument } from 'pdf-lib'
import { readInfo } from '@injoffice/pdf/browser'
import { rotateLastPage } from '../.examples-dist/pdf.js'
import { preparePdfRotation } from '../.examples-dist/pdf-agent.js'
import { createQuarterlyReview } from '../.examples-dist/presentation.js'
import { revenueChart, salesByRegion } from '../.examples-dist/spreadsheet-tools.js'

async function fixture() {
  const pdf = await PDFDocument.create()
  pdf.addPage([600, 800]).drawText('Quarterly operating review', { x: 40, y: 740 })
  pdf.addPage([600, 800]).drawText('Revenue and delivery', { x: 40, y: 740 })
  return pdf.save()
}
test('PDF sample rotates only the final page without changing its input', async () => {
  const bytes = await fixture()
  const original = bytes.slice()
  const { after, output } = await rotateLastPage(bytes)
  assert.deepEqual(bytes, original)
  assert.equal(after.pageCount, 2)
  assert.deepEqual(after.pages.map(page => page.rotation), [0, 90])
  assert.notDeepEqual(output, original)
})
test('PDF sample rejects malformed input', async () => {
  await assert.rejects(rotateLastPage(new Uint8Array([1, 2, 3])))
})
test('agent sample separates planning, exact-plan approval, and verified commit', async () => {
  const bytes = await fixture()
  const source = bytes.slice()
  const review = await preparePdfRotation(bytes)
  assert.deepEqual(bytes, source)
  assert.ok(review.preview && review.diff)
  await assert.rejects(review.commitFromReview('another-plan', 'commit-1'), /Approval/)
  const result = await review.commitFromReview(review.plan.changeSetId, 'commit-1')
  assert.equal(result.receipt.verification.verified, true)
  assert.deepEqual((await readInfo(result.bytes)).pages.map(page => page.rotation), [90, 0])
  const retry = await review.commitFromReview(review.plan.changeSetId, 'commit-1')
  assert.deepEqual(retry.bytes, result.bytes, 'retry must not rotate twice')
  assert.deepEqual(bytes, source)
})
test('authored presentation sample compiles a real native model', () => {
  const result = createQuarterlyReview()
  assert.equal(result.spec.slides.length, 1)
  assert.equal(result.nativeDeck.slides.length, 1)
  assert.match(JSON.stringify(result.nativeDeck), /Quarterly review/)
})
test('chart and pivot examples produce useful data', () => {
  assert.match(JSON.stringify(revenueChart()), /165/)
  const pivot = salesByRegion()
  assert.ok(pivot.grid.some(row => row[0] === 'West' && row[1] === 35))
  assert.ok(pivot.grid.some(row => row[0] === 'East' && row[1] === 35))
})
