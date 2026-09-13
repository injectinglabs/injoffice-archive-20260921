import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { expect, it } from 'vitest'
import { NativeRichTextSpans } from './NativeRichTextSpans'
import type { NativeRichTextCellV1 } from '../../../../packages/sheets/src/nativeRichTextPreviewV1'
it('resets each run to the cell font and safely preserves exact stored text', () => {
  const entry = { text: 'bold normal <&', status: 'available', runs: [
    { text: 'bold ', properties: 'direct', bold: true, omitted: [] },
    { text: 'normal <&', properties: 'direct', font_color: '#FF0000', omitted: ['font-family-hint'] },
  ] } as NativeRichTextCellV1
  const html = renderToStaticMarkup(createElement(NativeRichTextSpans, { entry, base: { font_name: 'Calibri', bold: false, italic: false }, normal: { font_name: 'Calibri' }, loadedFont: 'exact-calibri' }))
  expect(html).toContain('font-weight="700"'); expect(html).toContain('font-weight="400"')
  expect(html).toContain('font-family="exact-calibri"'); expect(html).toContain('normal &lt;&amp;')
  expect(html).toContain('white-space:pre'); expect(html).toContain('approximate cell-font fallback')
  expect(html).not.toContain('<script')
})
it('retains plaintext without spans for omitted styling', () => {
  const entry = { text: 'plain <text>', status: 'omitted' } as NativeRichTextCellV1
  expect(renderToStaticMarkup(createElement(NativeRichTextSpans, { entry, base: {} }))).toBe('plain &lt;text&gt;')
})
