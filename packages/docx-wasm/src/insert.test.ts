import {describe,it,expect} from 'vitest'
import {validateInsert} from './insert.js'
import type {NativeDocxDocumentV1} from '@injoffice/docs/native-docx'
const document = {body:{blocks:[{paragraph:{id:'p',anchor:{xml_sha256:'sha'},edit_policy:{allowed_operations:['page_break.insert','block.insert_after']},runs:[{id:'r',kind:'text',text:'A😀B'}]}}]}} as unknown as NativeDocxDocumentV1
const base = {target_kind:'paragraph',target_id:'p',expected_xml_sha256:'sha',operation:'page_break.insert',split:{run_id:'r',offset_utf16:3}}
describe('native insert boundary',()=>{
 it('carries a valid caret and refuses split surrogates, stale anchors, unknown fields and mixed batches',()=>{
  expect(validateInsert(document,[base]).mutations[0]).toEqual(base)
  for(const value of [{...base,split:{run_id:'r',offset_utf16:2}},{...base,expected_xml_sha256:'old'},{...base,text:''}]) expect(()=>validateInsert(document,[value])).toThrow()
  expect(()=>validateInsert(document,[base,base])).toThrow()
 })
 it('admits empty paragraphs but refuses nonempty text',()=>{
  const value={target_kind:'paragraph',target_id:'p',expected_xml_sha256:'sha',operation:'block.insert_after',text:''}
  expect(validateInsert(document,[value]).mutations[0]).toEqual(value)
  expect(()=>validateInsert(document,[{...value,text:'no'}])).toThrow()
 })
})
