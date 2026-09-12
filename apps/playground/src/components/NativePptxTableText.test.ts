import {createElement} from 'react'
import {renderToStaticMarkup} from 'react-dom/server'
import {describe,it,expect} from 'vitest'
import {readFileSync} from 'node:fs'
import type {NativePptxTableInspection} from '@injoffice/pptx-wasm'
import {NativePptxTableText,NativePptxTableTextResult} from './NativePptxTableText'

function fixture():NativePptxTableInspection{
 return {protocol:'pptx-table-content-inspection-v1',package_sha256:'a'.repeat(64),source_revision:`rev-${'a'.repeat(64)}`,omissions:[],tables:[{
  slide_id:'slide-1',slide_index:0,part_name:'ppt/slides/slide1.xml',part_sha256:'b'.repeat(64),slide_source_sha256:'c'.repeat(64),object_id:'cNvPr-15',source_sha256:'d'.repeat(64),
  rect:{x:20,y:30,width:200,height:100},warnings:['Authored borders and fonts omitted.'],cells:[
   {row:0,column:0,rect:{x:0,y:0,width:100,height:100},paragraphs:['First paragraph','<script>alert(1)</script>']},
   {row:0,column:1,rect:{x:100,y:0,width:100,height:100},paragraphs:['']},
  ],
 }]}
}
const render=(inspection=fixture())=>renderToStaticMarkup(createElement(NativePptxTableTextResult,{inspection}))
describe('read-only PPTX table source view',()=>{
 it('shows source cells and limits without pretending to render authored table geometry',()=>{
  const html=render()
  expect(html).toContain('1 table inspected; 0 omissions')
  expect(html).toContain('not PowerPoint rendering')
  expect(html).toContain('Slide 1, table 1')
  expect(html).toContain('data-source-cell="1:1"')
  expect(html).toContain('data-source-cell="1:2"')
  expect(html).toContain('Authored borders and fonts omitted.')
  expect(html).toContain('These measurements are not applied')
  expect(html).not.toContain('<svg')
  expect(html).not.toContain('contenteditable')
  expect(html).not.toContain('<input')
 })
 it('escapes source text, preserves paragraph order and marks empty paragraphs',()=>{
  const html=render()
  expect(html).toContain('&lt;script&gt;alert(1)&lt;/script&gt;')
  expect(html).not.toContain('<script>')
  expect(html.indexOf('First paragraph')).toBeLessThan(html.indexOf('&lt;script&gt;'))
  expect(html).toContain('[Empty paragraph]')
  expect(html).toContain('dir="auto"')
 })
 it('keeps empty results and source omissions visible',()=>{
  const inspection=fixture();inspection.tables=[];inspection.omissions=[{slide_id:'slide-1',object_id:'unavailable',reason:'Unqualified hidden source.'}]
  const html=render(inspection)
  expect(html).toContain('0 tables inspected; 1 omission')
  expect(html).toContain('No table text matches this inspection profile')
  expect(html).toContain('Unqualified hidden source.')
  expect(html).toContain('Omitted source content (1)')
 })
 it('requires an explicit browser action and does not upload source',()=>{
  const html=renderToStaticMarkup(createElement(NativePptxTableText,{bytes:new Uint8Array([1]),sourceRevision:'rev-source'}))
  expect(html).toContain('Show read-only table text')
  expect(html).not.toContain('Read-only presentation table text')
  const source=readFileSync(new URL('./NativePptxTableText.tsx',import.meta.url),'utf8')
  expect(source).toContain('client.inspectTables(bytes,{signal:controller.signal})')
  expect(source).toContain('client?.terminate()')
  expect(source).toContain('result?.bytes===bytes')
  expect(source).toContain('inspection.source_revision!==sourceRevision')
  expect(source).not.toMatch(/\bfetch\s*\(/)
 })
})
