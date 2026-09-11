import {describe,it,expect} from 'vitest'
import type {NativeElement,NativePptxDeck} from '@injoffice/pptx-native'
import {assertNativePptx} from '@injoffice/pptx-native'
import {quarterTurnTransform} from './geometry.js'

function deck(q:1|2|3):NativePptxDeck {
 const element:NativeElement={kind:'shape',id:'rotated',provenance:'authored',transform:{x:100,y:200,cx:400,cy:200,quarterTurns:q},preset:'triangle',fill:'123456',paragraphs:[],passthrough:[],compatibility:{status:'editable',diagnostics:[]}}
 return {contractVersion:'pptx-native/v1',documentId:'rotation',origin:'authored',size:{cx:1000,cy:1000},assets:[],compatibility:{status:'editable',diagnostics:[]},slides:[{id:'slide',provenance:'authored',elements:[element],passthrough:[],compatibility:{status:'editable',diagnostics:[]}}]}
}
describe('source quarter-turn affine',()=>{
 it.each([
  [1,{aPpm:0,bPpm:1000000,cPpm:-1000000,dPpm:0,txEmu:400,tyEmu:100}],
  [2,{aPpm:-1000000,bPpm:0,cPpm:0,dPpm:-1000000,txEmu:500,tyEmu:400}],
  [3,{aPpm:0,bPpm:-1000000,cPpm:1000000,dPpm:0,txEmu:200,tyEmu:500}],
 ] as const)('computes exact clockwise quarter turn %i around the frame center',(q,expected)=>{
  expect(quarterTurnTransform(deck(q).slides[0]!.elements[0]!.transform)).toEqual(expected)
  expect(()=>assertNativePptx(deck(q))).not.toThrow()
 })
 it('rejects fractional centers rather than rounding',async()=>{
  const value=deck(1);value.slides[0]!.elements[0]!.transform.cx=401
  expect(()=>quarterTurnTransform(value.slides[0]!.elements[0]!.transform)).toThrow('fractional EMU')
  expect(()=>assertNativePptx(value)).toThrow()
 })
 it('rejects source rotations on unmodeled group coordinate systems',async()=>{
  const value=deck(1),child=value.slides[0]!.elements[0]!
  value.slides[0]!.elements=[{kind:'group',id:'group',provenance:'authored',transform:{x:0,y:0,cx:1000,cy:1000,quarterTurns:1},children:[child],passthrough:[],compatibility:{status:'editable',diagnostics:[]}}]
  expect(()=>assertNativePptx(value)).toThrow()
 })
})
