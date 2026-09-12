import {describe,it,expect} from 'vitest'
import {compactNativeGeneralNumberPreviewV1 as compact} from './nativeCompactGeneralPreviewV1.js'
import {formatNativeSheetCellDisplayV2} from './nativeCellDisplayV2.js'
describe('explicit host compact General preview',()=>{
 it.each([
  ['5.3333333333333304','5.333333'],['4.3333333333333304','4.333333'],['11.3333333333333','11.33333'],
  ['-1.2345675','-1.234568'],['9.9999999','10'],['9999999.5','1e+7'],['-0.0000','0'],
  ['.00000012345675','1.234568e-7'],['1e-128','1e-128'],['1e128','1e+128'],['12E+0','12'],
  ['0.000001','0.000001'],['9007199254740993','9.007199e+15'],['+00012.0000','12'],
 ])('%s → %s',(source,text)=>expect(compact(source)).toMatchObject({status:'ready',text,policy:'host-general-seven-significant-v1'}))
 it('leaves ordinary integers unchanged and strict General lexical intact',()=>{
  for(let i=1;i<=12;i++)expect(compact(String(i))).toMatchObject({text:String(i)})
  expect(formatNativeSheetCellDisplayV2('number','5.3333333333333304','General')).toEqual({status:'ready',text:'5.3333333333333304'})
 })
 it.each(['NaN','Infinity','0x1p2','1e129','1e-129','1e0000','1,234','1/3','1'.repeat(129)])('refuses bounded/unknown lexical %s',value=>expect(compact(value).status).toBe('refused'))
})
