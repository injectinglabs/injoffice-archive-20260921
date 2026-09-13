import {describe,it,expect} from 'vitest'
import {createElement} from 'react'
import {renderToStaticMarkup} from 'react-dom/server'
import {PptxFilePreviewVector} from './components/PptxFilePreview'
import type {NativeElement,NativePptxDeck} from '@injoffice/pptx-native'
import {compileFilePreviewGeometry,filePreviewPath} from './filePreviewGeometry'
const compatibility={status:'preserveOnly' as const,diagnostics:[{severity:'warning' as const,code:'fixture.preview',message:'Preview only'}]}
const body={leftInsetEmu:12700,rightInsetEmu:25400,topInsetEmu:38100,bottomInsetEmu:50800,wrap:'none' as const,verticalAnchor:'top' as const,autoFit:'none' as const,horizontalOverflow:'overflow' as const,verticalOverflow:'overflow' as const}
export function previewShape():Extract<NativeElement,{kind:'shape'}>{return {kind:'shape',id:'shape',provenance:'authored',transform:{x:0,y:0,cx:2000000,cy:2000000,rotationAngle:5400000},preset:'triangle',fill:'336699',paragraphs:[{runs:[{text:'Local browser text'}]}],textBody:body,passthrough:[],compatibility}}
export function previewDeck(element:NativeElement):NativePptxDeck{return {contractVersion:'pptx-native/v1',documentId:'preview',origin:'authored',size:{cx:12000000,cy:9000000},assets:[],compatibility,slides:[{id:'slide',provenance:'authored',elements:[element],passthrough:[],compatibility}]}}
describe('public local-file geometry projection',()=>{
 it('keeps exact source hierarchy and original text immutable without invoking fonts',async()=>{
  const child=previewShape(),group:NativeElement={kind:'group',id:'group',provenance:'authored',transform:{x:0,y:0,cx:4000000,cy:2000000,rotationAngle:0},childTransform:{x:0,y:0,cx:2000000,cy:2000000},children:[child],passthrough:[],compatibility}
  const deck=previewDeck(group),before=JSON.stringify(deck),geometry=await compileFilePreviewGeometry(deck,0),node=geometry.tree.nodes[0]!
  if(node.kind!=='group'||node.children[0]?.kind!=='shape')throw Error('group/shape')
  expect(geometry.matrices.get(node.transform)).toEqual([1,0,0,1,0,0])
  expect(geometry.matrices.get(node.children[0].transform)).toEqual([0,2,-1,0,3000000/12700,-1000000/12700])
  expect(node.children[0].textBody?.paragraphs).toEqual([])
  expect(geometry.originals.get('shape')).toBe(child);expect(JSON.stringify(deck)).toBe(before)
 })
 it('uses public compound paint including curves/arcs/tones and outside-frame control points',async()=>{
  const shape=previewShape();delete shape.preset;shape.transform.rotationAngle=0;shape.fill='000000'
  shape.geometry={profile:'drawingml-paths-v1',textRect:{x:0,y:0,cx:2000000,cy:2000000},paths:[{fillMode:'lighten',stroke:false,commands:[{kind:'moveTo',x:0,y:0},{kind:'quadBezierTo',x1:-12700,y1:25400,x:25400,y:25400},{kind:'cubicBezierTo',x1:38100,y1:25400,x2:50800,y2:25400,x:50800,y:38100},{kind:'arcTo',rx:12700,ry:12700,x:25400,y:38100,largeArc:false,clockwise:true},{kind:'close'}]},{fillMode:'none',stroke:false,commands:[{kind:'moveTo',x:0,y:0},{kind:'lineTo',x:12700,y:12700}]}]}
  const geometry=await compileFilePreviewGeometry(previewDeck(shape),0),node=geometry.tree.nodes[0]!
  expect(node.clip).toBeUndefined();const paths=geometry.paths.get('shape')!
  expect(paths).toHaveLength(2);expect(paths[0]!.fill).toBe('AAAAAA');expect(paths[1]!.fill).toBeUndefined()
  expect(filePreviewPath(paths[0]!.path)).toBe('M 0 0 Q -1 2 2 2 C 3 2 4 2 4 3 A 1 1 0 0 1 2 3 Z')
 })
 it('keeps refused geometry unavailable and rejects malformed source before projection',async()=>{
  const shape=previewShape();shape.compatibility={status:'refused',diagnostics:[{severity:'refusal',code:'fixture.unsupported',message:'unsupported visual clause'}]}
  const input=previewDeck(shape);input.compatibility=shape.compatibility;input.slides[0]!.compatibility=shape.compatibility
  const geometry=await compileFilePreviewGeometry(input,0)
  expect(geometry.tree.nodes[0]?.kind).toBe('placeholder');expect(geometry.paths.size).toBe(0)
  shape.transform.rotationAngle=-1;await expect(compileFilePreviewGeometry(previewDeck(shape),0)).rejects.toThrow('validated')
 })
 it('retains text orientation and insets for empty projection without pretending font qualification',async()=>{
  const shape=previewShape();shape.transform={...shape.transform,rotationAngle:1800000,flipH:true}
  const geometry=await compileFilePreviewGeometry(previewDeck(shape),0),node=geometry.tree.nodes[0]!
  if(node.kind!=='shape'||!node.textBody)throw Error('body')
  expect(node.textBody.bounds).toEqual({x:512700,y:1038100,cx:961900,cy:911100})
  expect(geometry.matrices.get(node.textBody.orientationTransform!)).toEqual([-1,0,0,1,2000000/12700,0])
 })
 it('renders every tone independently, honors path stroke flags and never clips a callout to its frame',async()=>{
  const shape=previewShape();delete shape.preset;shape.transform.rotationAngle=0;shape.fill='000000';shape.stroke={color:'FF0000',widthEmu:12700,cap:'round',join:'round',dash:'solid'}
  shape.geometry={profile:'drawingml-paths-v1',textRect:{x:0,y:0,cx:2000000,cy:2000000},paths:(['norm','none','darken','darkenLess','lighten','lightenLess'] as const).map((fillMode,i)=>({fillMode,stroke:i%2===0,commands:[{kind:'moveTo' as const,x:-12700,y:i*12700},{kind:'lineTo' as const,x:25400,y:i*12700},{kind:'lineTo' as const,x:25400,y:(i+1)*12700},{kind:'close' as const}]}))}
  const deck=previewDeck(shape),geometry=await compileFilePreviewGeometry(deck,0),html=renderToStaticMarkup(createElement(PptxFilePreviewVector,{deck,geometry}))
  expect(html.match(/<path /g)).toHaveLength(6);expect(html).toContain('M -1 0');expect(html).not.toContain('clipPath');expect(html.match(/stroke="#FF0000"/g)).toHaveLength(3);expect(html).toContain('fill="#AAAAAA"');expect(html).toContain('fill="#7C7C7C"');expect(html).toContain('fill="none"')
 })
 it('bounds nested object traversal at500 while preserving source membership and order',async()=>{
  const children=Array.from({length:501},(_,i)=>{const shape=previewShape();shape.id=`shape-${i}`;shape.transform.rotationAngle=0;shape.paragraphs=[];delete shape.textBody;return shape})
  const group:NativeElement={kind:'group',id:'group',provenance:'authored',transform:{x:0,y:0,cx:4000000,cy:4000000},children,passthrough:[],compatibility}
  const deck=previewDeck(group),geometry=await compileFilePreviewGeometry(deck,0),node=geometry.tree.nodes[0]!
  if(node.kind!=='group')throw Error('group')
  expect(node.children).toHaveLength(499);expect(geometry.omitted).toBe(2);expect(node.children.at(-1)?.sourceElementId).toBe('shape-498');expect(children).toHaveLength(501)
 })
 it('retains unsupported body qualification instead of empty-run promotion',async()=>{
  const shape=previewShape();shape.textBody={...body,writingMode:'vertical-clockwise'};shape.paragraphs=[{bullet:false,level:0,runs:[{text:'漢字'}]}]
  const deck=previewDeck(shape),geometry=await compileFilePreviewGeometry(deck,0),node=geometry.tree.nodes[0]!
  if(node.kind!=='shape')throw Error('shape')
  expect(geometry.textIssues.get(shape.id)).toContain('Vertical text');const html=renderToStaticMarkup(createElement(PptxFilePreviewVector,{deck,geometry}))
  expect(html).toContain('Text preview unavailable');expect(html).not.toContain('漢字')
 })

})
