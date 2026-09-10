import { applyPageOps, readInfo } from '@injoffice/pdf/browser'

/** Node or browser: receive bytes from your host, return new bytes. */
export async function rotateLastPage(source: Uint8Array) {
  const before = await readInfo(source)
  if (before.pageCount === 0) throw new Error('The PDF has no pages')
  // PDF page selectors are one-based, unlike spreadsheet coordinates.
  const output = await applyPageOps(source, [
    { type: 'rotate', pages: [before.pageCount], degrees: 90 },
  ])
  const after = await readInfo(output)
  if (after.pageCount !== before.pageCount) throw new Error('Page count changed')
  return { output, before, after }
}
