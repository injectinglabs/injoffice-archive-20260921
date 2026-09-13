import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { expect, it } from 'vitest'
import { NativeRichTextSpans } from './NativeRichTextSpans'
import type { NativeRichTextCellV1 } from '@injoffice/sheets/browser'
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

it('exposes a clearly labeled source sample and all per-run omissions without page geometry', async () => {
  const { NativeRichTextDetails } = await import('./NativeRichTextSpans')
  const entry = { sheet_id: '1', sheet_part: 'xl/worksheets/sheet1.xml', ref: 'A1', style_id: 0, source_part: 'xl/sharedStrings.xml', shared_index: '0', storage: 'shared', text: 'sample', status: 'available', warnings: ['Approximate display.'], runs: [{ text: 'sample', properties: 'direct', font_color: '#FF0000', omitted: ['baseline', 'font-family-hint'] }] } as NativeRichTextCellV1
  const html = renderToStaticMarkup(createElement(NativeRichTextDetails, { entries: [entry], styles: [{ id: 0, effective: { font_name: 'Calibri' } }], warnings: [] }))
  expect(html).toContain('not worksheet positions or page geometry'); expect(html).toContain('shared string 0'); expect(html).toContain('Source rich-text sample A1'); expect(html).toContain('font-family-hint'); expect(html).toContain('Stored text: sample')
})

it('decorates individual runs and resets absent and explicit-none runs', () => {
 const entry = { text: 'on absent off', status: 'available', runs: [
  {text:'on ',properties:'direct',underline:'single',underline_origin:'default-val',omitted:[]},
  {text:'absent ',properties:'direct',omitted:[]},
  {text:'off',properties:'direct',underline:'none',underline_origin:'explicit-val',omitted:[]},
 ] } as NativeRichTextCellV1
 const html = renderToStaticMarkup(createElement(NativeRichTextSpans, {entry,base:{}}))
 expect(html.match(/text-decoration:underline solid/g)).toHaveLength(1)
 expect(html.match(/text-decoration:none/g)).toHaveLength(2)
 expect(html).toContain('schema-default single'); expect(html).toContain('undecorated host fallback')
})

it('paints resolved theme faces, discloses the declared name and resets following direct runs', () => {
 const entry = {text:'Major plain',status:'available',runs:[
  {text:'Major ',properties:'direct',font_name:'Cambria',font_scheme:'major',declared_font_name:'Arial',theme_part:'xl/theme/theme1.xml',theme_sha256:`sha256:${'b'.repeat(64)}`,omitted:[]},
  {text:'plain',properties:'direct',font_name:'Calibri',omitted:[]},
 ]} as NativeRichTextCellV1
 const html=renderToStaticMarkup(createElement(NativeRichTextSpans,{entry,base:{font_name:'Courier'}}))
 expect(html).toContain('font-family="Cambria"');expect(html).toContain('font-family="Calibri"')
 expect(html).not.toContain('font-family="Arial"');expect(html).toContain('declared Arial, resolved theme Latin face Cambria')
})
