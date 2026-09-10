import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it } from 'vitest'
import { ParagraphView, TableView } from './DocsPage'
import type { NativeDocxParagraphV1, NativeDocxTableV1 } from '../../../../packages/docs/src/nativeContract'

const anchor = { part_name: 'word/document.xml', path: '/p', start_byte: 1, end_byte: 2, xml_sha256: `sha256:${'a'.repeat(64)}` }
const edit_policy = { mode: 'read-only' as const, allowed_operations: [] }
const paragraph: NativeDocxParagraphV1 = { id: 'p', anchor, edit_policy, properties: {}, runs: [] }

describe('DOCX content preview markup', () => {
  it('shows ordinary tabs and breaks without injecting debug text and applies highlighting', () => {
    const html = renderToStaticMarkup(createElement(ParagraphView, { paragraph: { ...paragraph, runs: [
      { id: 'text', anchor, kind: 'text', text: '<safe>', properties: { highlight: 'darkYellow', underline: 'double' } },
      { id: 'tab', anchor, kind: 'control', control: 'tab' },
      { id: 'break', anchor, kind: 'control', control: 'line-break' },
    ] } }))
    expect(html).toContain('&lt;safe&gt;')
    expect(html).toContain('background-color:#808000')
    expect(html).toContain('text-decoration-style:double')
    expect(html).toContain('\t')
    expect(html).toContain('<br/>')
    expect(html).not.toContain('[line-break]')
  })
  it('labels unresolved lists rather than inventing a bullet', () => {
    const html = renderToStaticMarkup(createElement(ParagraphView, { paragraph: { ...paragraph, properties: { numbering: { num_id: '1', level: 0 } } } }))
    expect(html).toContain('[list]')
    expect(html).not.toContain('•')
  })
  it('emits real row spans and column headers with authored borders', () => {
    const table: NativeDocxTableV1 = { id: 'table', anchor, edit_policy, borders: { top: { style: 'none', size_eighth_points: 0 } }, rows: [
      { id: 'row1', anchor, repeat_header: true, cells: [{ id: 'a', anchor, vertical_merge: 'restart', grid_span: 2, paragraphs: [paragraph] }] },
      { id: 'row2', anchor, repeat_header: false, cells: [{ id: 'b', anchor, vertical_merge: 'continue', grid_span: 2, paragraphs: [paragraph] }] },
    ] }
    const html = renderToStaticMarkup(createElement(TableView, { table }))
    expect(html).toContain('<th scope="col" colSpan="2" rowSpan="2"')
    expect(html).toContain('border-top:none')
    expect(html).not.toContain('continued merged cell')
  })
  it('keeps white source text but visibly explains a readability outline when its background is unavailable', () => {
    const table: NativeDocxTableV1 = { id: 'table', anchor, edit_policy, rows: [{ id: 'row', anchor, repeat_header: true, cells: [{ id: 'cell', anchor, vertical_merge: 'none', grid_span: 1, paragraphs: [{ ...paragraph, runs: [{ id: 'white', anchor, kind: 'text', text: 'Visible heading', properties: { color: 'FFFFFF' } }] }] }] }] }
    const html = renderToStaticMarkup(createElement(TableView, { table }))
    expect(html).toContain('docx-missing-background')
    expect(html).toContain('Background unavailable')
    expect(html).toContain('color:#FFFFFF')
    table.rows[0].cells[0].shading_rgb = '234F78'
    expect(renderToStaticMarkup(createElement(TableView, { table }))).not.toContain('Background unavailable')
  })
})
