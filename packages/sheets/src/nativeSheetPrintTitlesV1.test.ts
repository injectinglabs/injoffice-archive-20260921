import {describe,it,expect} from 'vitest'
import {decodeNativeSheetPrintTitlesV1} from './nativeSheetPrintTitlesV1.js'
describe('closed saved print titles contract',()=>{
 const valid=()=>({sheet_id:'7',sheet_part:'xl/worksheets/sheet1.xml',status:'available',warnings:['Source'],rows:{start:0,end:1},columns:{start:0,end:2}})
 it('snapshots both axes and unavailable entries',()=>{
  const input=valid(),out=decodeNativeSheetPrintTitlesV1([input]);expect(out).toEqual([input]);expect(out[0]).not.toBe(input)
  expect(decodeNativeSheetPrintTitlesV1([{sheet_id:'7',sheet_part:'xl/a.xml',status:'unavailable',warnings:['Missing']}])[0]!.status).toBe('unavailable')
 })
 it('rejects extras, duplicate identity, missing axes, malformed bounds and executable data',()=>{
  const bad:unknown[]=[null,{},[{...valid(),extra:1}],[valid(),valid()],[{...valid(),status:'unavailable'}],[{sheet_id:'7',sheet_part:'xl/a.xml',status:'available',warnings:['Source']}]]
  for(const range of [{start:-0,end:1},{start:-1,end:1},{start:2,end:1},{start:0,end:1048576},{start:0,end:1.5},{start:0,end:1,extra:1},null])bad.push([{...valid(),rows:range}])
  bad.push([{...valid(),columns:{start:0,end:16384}}],[{...valid(),sheet_part:'../a.xml'}],[{...valid(),warnings:['bad\nwarning']}])
  let called=false;bad.push([Object.defineProperty(valid(),'rows',{enumerable:true,get(){called=true;return {start:0,end:0}}})])
  for(const input of bad)expect(()=>decodeNativeSheetPrintTitlesV1(input)).toThrow()
  expect(called).toBe(false)
 })
})
