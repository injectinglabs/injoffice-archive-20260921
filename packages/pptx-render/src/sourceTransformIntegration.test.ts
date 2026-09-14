import {expect,it} from 'vitest'
import type {NativePptxDeck,NativeElement} from '@injoffice/pptx-native'
import {validateNativePptx} from '@injoffice/pptx-native'
import {compileNativePptxSlide,createRecordingPaintSurface,paintSlideRenderTree,renderTransformMatrix,SourceAffineBudget,type NativePptxTextLayout} from './index.js'
const layout:NativePptxTextLayout={manifest:{version:1,manifestId:'geometry',revision:'1',faces:[{faceId:'unused',family:'Unused',weight:400,style:'normal',stretch:100,source:{kind:'host',resourceId:'unused',contentDigest:`sha256:${'0'.repeat(64)}`}}],fallbackChains:[]},resolver:{providerId:'unused',providerRevision:'1',resolve(){throw Error('No text')},load(){throw Error('No text')}},shaper:{providerId:'unused',providerRevision:'1',shape(){throw Error('No text')}},defaults:{fontFamilies:['Unused'],fontSizeHundredthPt:1200,script:'Latn',language:'en-US',direction:'ltr'}}
const compatibility={status:'preserveOnly' as const,diagnostics:[{severity:'warning' as const,code:'test.preview',message:'Preview-only affine'}]}
function shape():Extract<NativeElement,{kind:'shape'}>{return {kind:'shape',id:'shape',provenance:'authored',transform:{x:0,y:0,cx:2000000,cy:2000000,rotationAngle:5400000},preset:'triangle',fill:'336699',paragraphs:[],passthrough:[],compatibility}}
function deck(element:NativeElement):NativePptxDeck{return {contractVersion:'pptx-native/v1',documentId:'affine',origin:'authored',size:{cx:12000000,cy:9000000},assets:[],compatibility,slides:[{id:'slide',provenance:'authored',elements:[element],passthrough:[],compatibility}]}}
const matrix=(t:Parameters<typeof renderTransformMatrix>[0])=>renderTransformMatrix(t,new SourceAffineBudget())
it('uses Annex own-axis scaling for rotated leaves in a nonuniform source-mode group',async()=>{
 const child=shape(),group:NativeElement={kind:'group',id:'group',provenance:'authored',transform:{x:0,y:0,cx:4000000,cy:2000000,rotationAngle:0},childTransform:{x:0,y:0,cx:2000000,cy:2000000},children:[child],passthrough:[],compatibility}
 const input=deck(group),before=JSON.stringify(input),tree=await compileNativePptxSlide(input,0,{textLayout:layout}),node=tree.nodes[0]!
 expect(node.kind).toBe('group');if(node.kind!=='group')throw Error('group')
 expect(matrix(node.transform)).toEqual([1,0,0,1,0,0])
 expect(matrix(node.children[0]!.transform)).toEqual([0,2,-1,0,3000000,-1000000])
 expect(JSON.stringify(input)).toBe(before)
})
it('retains authored legacy conventional group and quarter-turn behavior',async()=>{
 const child=shape();delete child.transform.rotationAngle;child.transform.quarterTurns=1
 const group:NativeElement={kind:'group',id:'group',provenance:'authored',transform:{x:0,y:0,cx:4000000,cy:2000000},childTransform:{x:0,y:0,cx:2000000,cy:2000000},children:[child],passthrough:[],compatibility}
 const node=(await compileNativePptxSlide(deck(group),0,{textLayout:layout})).nodes[0]!
 if(node.kind!=='group')throw Error('group')
 expect(matrix(node.transform)).toEqual([2,0,0,1,0,0]);expect(matrix(node.children[0]!.transform)).toEqual([0,1,-1,0,2000000,0])
})
it('transports a noncardinal reflection without PPM quantization through public paint',async()=>{
 const input=shape();input.transform={x:1000000,y:2000000,cx:4000000,cy:2000000,rotationAngle:1800000,flipH:true}
 const tree=await compileNativePptxSlide(deck(input),0,{textLayout:layout}),surface=createRecordingPaintSurface();paintSlideRenderTree(tree,surface)
 const transform=surface.finish().find(c=>c.kind==='transform');if(transform?.kind!=='transform')throw Error('transform')
 expect(transform.transform.sourceAffine?.profile).toBe('rational-affine-v1')
 const m=matrix(transform.transform)
 expect(m[0]).toBeCloseTo(-Math.sqrt(3)/2,14);expect(m[1]).toBeCloseTo(-.5,14);expect(m[2]).toBeCloseTo(-.5,14);expect(m[3]).toBeCloseTo(Math.sqrt(3)/2,14)
 expect(m[4]).toBeCloseTo(3500000+1000000*Math.sqrt(3),7)
 expect(m[5]).toBeCloseTo(4000000-500000*Math.sqrt(3),7)
})
it('rejects ambiguous, editable, and noncanonical native affine requests',()=>{
 for(const transform of [{rotationAngle:1,quarterTurns:1},{rotationAngle:-1},{rotationAngle:21600000},{rotationAngle:-0},{flipH:'true'}]){
  const value=shape();Object.assign(value.transform,transform);expect(validateNativePptx(deck(value)).ok).toBe(false)
 }
 const value=shape();value.compatibility={status:'editable',diagnostics:[]};expect(validateNativePptx(deck(value)).ok).toBe(false)
})
it('rejects uncertain geometry outside the frame instead of clipping it to hide overflow',async()=>{
 const value=shape();delete value.preset;value.transform.rotationAngle=1800000
 value.geometry={profile:'drawingml-paths-v1',textRect:{x:0,y:0,cx:100,cy:100},paths:[{fillMode:'norm',stroke:false,commands:[{kind:'moveTo',x:0,y:0},{kind:'quadBezierTo',x1:10000000000000,y1:0,x:1,y:1}]}]}
 await expect(compileNativePptxSlide(deck(value),0,{textLayout:layout})).rejects.toThrow(/uncertainty/)
})
it('rejects a rational transform combined with a second PPM operation',async()=>{
 const value=shape();value.transform.rotationAngle=1
 const transform=(await compileNativePptxSlide(deck(value),0,{textLayout:layout})).nodes[0]!.transform
 expect(()=>matrix({...transform,txEmu:1})).toThrow('cannot also')
})

it('compiles 200 ordinary noncardinal cubic callouts with one aggregate affine budget',async()=>{
 const input=deck(shape())
 input.slides[0]!.elements=Array.from({length:200},(_,i)=>{
  const value=shape();value.id=`shape-${i}`;value.transform={x:(i%20)*400000,y:Math.floor(i/20)*400000,cx:300000,cy:200000,rotationAngle:1800000+i*6000};delete value.preset
  value.geometry={profile:'drawingml-paths-v1',textRect:{x:0,y:0,cx:300000,cy:200000},paths:[{fillMode:'norm',stroke:false,commands:[{kind:'moveTo',x:0,y:0},{kind:'cubicBezierTo',x1:-100000,y1:200000,x2:400000,y2:-100000,x:300000,y:200000},{kind:'close'}]}]}
  return value
 })
 const tree=await compileNativePptxSlide(input,0,{textLayout:layout})
 expect(tree.nodes).toHaveLength(200);expect(tree.nodes.every(node=>node.transform.sourceAffine)).toBe(true)
 input.slides[0]!.elements=Array.from({length:1000},(_,i)=>({...input.slides[0]!.elements[i%200]!,id:`over-${i}`}))
 await expect(compileNativePptxSlide(input,0,{textLayout:layout})).rejects.toThrow('Aggregate affine operation')
})

it('does not reinterpret authored conventional groups when new text orientation is requested',async()=>{
 const child=shape();delete child.transform.rotationAngle
 child.textBody={leftInsetEmu:0,rightInsetEmu:0,topInsetEmu:0,bottomInsetEmu:0,wrap:'none',verticalAnchor:'top',autoFit:'none',horizontalOverflow:'overflow',verticalOverflow:'overflow',upright:true}
 const group:NativeElement={kind:'group',id:'group',provenance:'authored',transform:{x:0,y:0,cx:4000000,cy:2000000},childTransform:{x:0,y:0,cx:2000000,cy:2000000},children:[child],passthrough:[],compatibility}
 await expect(compileNativePptxSlide(deck(group),0,{textLayout:layout})).rejects.toThrow('authored conventional affine group')
 delete child.textBody.upright
 expect((await compileNativePptxSlide(deck(group),0,{textLayout:layout})).nodes).toHaveLength(1)
})
