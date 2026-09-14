import { readFileSync } from 'node:fs'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it } from 'vitest'
import { decodeXlsxSourceStylePreviewV1 } from '@injoffice/xlsx-wasm'
import { SourceStyleGrid, sourceStyleCellText } from './XlsxSourceStylePreview'
const raw = readFileSync(new URL('../../../../packages/xlsx-wasm/testdata/source-style-preview.json', import.meta.url), 'utf8')
const source = decodeXlsxSourceStylePreviewV1(raw, JSON.parse(raw).package_sha256)
describe('source-style read-only grid', () => {
  it('renders one merge anchor, source colors, cached display, and no edit/download controls', () => {
    const html = renderToStaticMarkup(<SourceStyleGrid preview={source} name="source.xlsx" />)
    expect(html).toContain('colSpan="2"')
    expect(html).toContain('background:#1E2761')
    expect(html).toContain('color:#FFFFFF')
    expect(html).toContain('3.00 €')
    expect(html).toContain('SUM(1,2)')
    expect(html).toContain('absent applyFill')
    expect(html).not.toContain('data-source-cell="B1"')
    expect(html).not.toMatch(/contenteditable|<input|download=/i)
  })
  it('keeps date formatting and explicitly reports unformatted fallback', () => {
    const cell = {...source.sheets[0]!.cells[2]!, lexical:'45292', formula:'',cached:false}
    const style = {...source.styles[1]!,number_format:'yyyy\\-mm\\-dd'}
    expect(sourceStyleCellText(cell,style,false)).toEqual({text:'2024-01-01'})
    const refused = sourceStyleCellText(cell,{...style,number_format:'[Unknown]'},false)
    expect(refused.text).toBe('45292');expect(refused.warning).toContain('saved value')
  })
})
