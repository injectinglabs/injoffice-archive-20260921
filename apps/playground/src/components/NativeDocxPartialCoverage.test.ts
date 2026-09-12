import {readFileSync} from 'node:fs'
import {describe,it,expect} from 'vitest'
import {renderToStaticMarkup} from 'react-dom/server'
import {createElement} from 'react'
import {createNativeDocxPartialContentPreviewV1} from '@injoffice/docs/native-docx'
import {NativeDocxPartialCoverage} from './NativeDocxPartialCoverage'
describe('partial DOCX inventory integration',()=>{
 it('renders validated library coverage separately from existing editor authority',()=>{
  const source=JSON.parse(readFileSync(new URL('../../../../testdata/docx-native/document-v1.json',import.meta.url),'utf8'))
  const coverage=createNativeDocxPartialContentPreviewV1(source,{policy:'source-text-with-omissions-v1',read_only:true})
  const html=renderToStaticMarkup(createElement(NativeDocxPartialCoverage,{coverage}))
  expect(html).toContain(`${coverage.coverage.visited_body_blocks} of 2 body blocks inspected`)
  expect(html).toContain('not a claim that the editor below omits the same content')
  expect(html).toContain('Read-only projection limitations')
  expect(html).not.toContain('contenteditable');expect(html).not.toContain('<button')
  expect(coverage.coverage.projected_text_runs).toBe(0)
 })
 it('escapes source identifiers and discloses UI inventory truncation',()=>{
  const source=JSON.parse(readFileSync(new URL('../../../../testdata/docx-native/document-v1.json',import.meta.url),'utf8'))
  const coverage=createNativeDocxPartialContentPreviewV1(source,{policy:'source-text-with-omissions-v1',read_only:true})
  coverage.omissions=Array.from({length:25},()=>({...coverage.omissions[0]!,source:{...coverage.omissions[0]!.source,scope_id:'<script>hostile</script>'}}))
  const html=renderToStaticMarkup(createElement(NativeDocxPartialCoverage,{coverage}))
  expect(html).not.toContain('<script>');expect(html).toContain('&lt;script&gt;')
  expect(html).toContain('5 additional source-bound limitations')
 })
})
