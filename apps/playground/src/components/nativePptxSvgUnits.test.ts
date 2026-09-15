import {describe,it,expect} from 'vitest'
import {nativePptxSvgNode,nativePptxSvgPath} from './nativePptxSvgUnits'
import type {PreviewNode} from '../../../pptx-page-paint-worker/src/contract'
describe('native SVG change of user units',()=>{
 it('rescales all path lengths while preserving arc angle and flags',()=>{
  expect(nativePptxSvgPath('M 12700 25400 L 38100 50800 Q 12700 25400 38100 50800 C 12700 25400 38100 50800 63500 76200 A 25400 38100 0 0 1 50800 63500 Z')).toBe('M 1 2 L 3 4 Q 1 2 3 4 C 1 2 3 4 5 6 A 2 3 0 0 1 4 5 Z')
 })
 it('preserves nested composed points, glyph scale, active rounded clips and strokes without mutating transport',()=>{
  const source:PreviewNode={kind:'group',transform:[0,2,-3,0,127000,254000],clip:{x:12700,y:25400,cx:1270000,cy:635000,radius:12700},children:[{kind:'group',transform:[.0127,0,0,-.0127,38100,50800],children:[{kind:'path',d:'M 12700 25400 L 38100 50800',fill:'FFFFFF',stroke:'000000',strokeWidth:25400,strokeMiterlimit:4}]}]}
  const before=JSON.stringify(source),normalized=nativePptxSvgNode(source)
  if(normalized.kind!=='group'||normalized.children[0]?.kind!=='group')throw Error('group')
  const glyph=normalized.children[0],path=glyph.children[0]
  expect(normalized.transform).toEqual([0,2,-3,0,10,20]);expect(glyph.transform).toEqual([.0127,0,0,-.0127,3,4])
  expect(normalized.clip).toEqual({x:1,y:2,cx:100,cy:50,radius:1})
  expect(path).toMatchObject({d:'M 1 2 L 3 4',strokeWidth:2,strokeMiterlimit:4})
  const originalPoint=[.0127*12700+38100,-.0127*25400+50800]
  expect(-3*(.0127*-2+4)+10).toBeCloseTo((-3*originalPoint[1]+127000)/12700,12)
  expect(2*(.0127*1+3)+20).toBeCloseTo((2*originalPoint[0]+254000)/12700,12)
  expect(JSON.stringify(source)).toBe(before)
 })
 it('rescales path clips with the same rule as painted paths',()=>{
  const source:PreviewNode={kind:'group',transform:[1,0,0,1,0,0],clip:{x:0,y:0,cx:1270000,cy:635000,d:'M 0 317500 A 635000 317500 0 0 1 635000 0 Z'},children:[]}
  const before=JSON.stringify(source),normalized=nativePptxSvgNode(source)
  if(normalized.kind!=='group')throw Error('group')
  expect(normalized.clip).toEqual({x:0,y:0,cx:100,cy:50,d:'M 0 25 A 50 25 0 0 1 50 0 Z'})
  expect(JSON.stringify(source)).toBe(before)
 })
 it('rescales rectangles, ellipses, placeholders and image frames while preserving crop percentages',()=>{
  for(const kind of ['rect','ellipse','placeholder','image'] as const){
   const source={kind,rect:{x:12700,y:25400,cx:38100,cy:50800},radius:12700,fill:'FFFFFF',strokeWidth:12700,label:'x',resourceId:'r',crop:{left:1000,top:2000,right:3000,bottom:4000}} as PreviewNode
   const output=nativePptxSvgNode(source)
   expect(output).toMatchObject({rect:{x:1,y:2,cx:3,cy:4}})
   if(output.kind==='image')expect(output.crop).toEqual({left:1000,top:2000,right:3000,bottom:4000})
   if(output.kind==='rect')expect(output.radius).toBe(1)
  }
 })
})
