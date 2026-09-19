import {describe,it,expect} from 'vitest'
import {parseNativeSheetHeaderFooterV1} from './nativeSheetHeaderFooterV1.js'

const facts={sheet_name:'Sheet1',page_number:2,page_count:5,default_font_size_points:10}

describe('authored header and footer format codes',()=>{
 it('resolves the section switches, the worksheet name and the page numbers',()=>{
  const sections=parseNativeSheetHeaderFooterV1('&LLeft&C&A&RPage &P of &N',facts)
  expect(sections?.map(section=>[section.align,section.runs.map(run=>run.text).join('')])).toEqual([
   ['left','Left'],['center','Sheet1'],['right','Page 2 of 5'],
  ])
 })
 it('keeps an uncoded run on the workbook Normal font at the authored default size',()=>{
  const runs=parseNativeSheetHeaderFooterV1('&C&A',facts)?.[0]!.runs
  expect(runs).toEqual([{text:'Sheet1',font_size_points:10,bold:false,italic:false}])
  expect(runs?.[0]!.font_name).toBeUndefined()
 })
 it('applies an authored face, style and size to the runs that follow it',()=>{
  const sections=parseNativeSheetHeaderFooterV1('&C&"Times New Roman,Bold Italic"&12&A',facts)
  expect(sections?.[0]!.runs).toEqual([{text:'Sheet1',font_name:'Times New Roman',font_size_points:12,bold:true,italic:true}])
 })
 it('reads a multi-digit size and an escaped ampersand as text',()=>{
  const sections=parseNativeSheetHeaderFooterV1('&C&14Q&&A',facts)
  expect(sections?.[0]!.runs).toEqual([{text:'Q&A',font_size_points:14,bold:false,italic:false}])
 })
 it('drops a section the source leaves empty rather than printing a blank run',()=>{
  expect(parseNativeSheetHeaderFooterV1('&L&C&A&R',facts)?.map(section=>section.align)).toEqual(['center'])
 })
 it('refuses the whole string for a code whose printed result it does not hold',()=>{
  // A dropped date or file path would print a header that is quietly wrong;
  // refusing leaves the caller saying the header is not painted at all.
  for(const code of ['&C&D','&C&T','&C&F','&C&Z','&C&G','&C&K0000FF Blue','&C&A&','&C&"Arial"','&C&"Arial,Outline"','&C&0A','&C&500A']){
   expect(parseNativeSheetHeaderFooterV1(code,facts),code).toBeUndefined()
  }
 })
 it('refuses a string, a face or a default size beyond its bounds',()=>{
  expect(parseNativeSheetHeaderFooterV1('&C'+'x'.repeat(1024),facts)).toBeUndefined()
  expect(parseNativeSheetHeaderFooterV1(`&C&"${'f'.repeat(129)},Regular"x`,facts)).toBeUndefined()
  expect(parseNativeSheetHeaderFooterV1('&C&A',{...facts,default_font_size_points:0})).toBeUndefined()
  expect(parseNativeSheetHeaderFooterV1('',facts)).toBeUndefined()
 })
})
