import {readFileSync} from 'node:fs'
import {resolve} from 'node:path'
import {expect,it} from 'vitest'
import {validateNativePptx,type NativePptxDeck,type NativeEvaluatedGeometry} from '@injoffice/pptx-native'
import {compileNativePptxSlide,createRecordingPaintSurface,paintSlideRenderTree,type NativePptxTextLayout} from './index.js'
const layout:NativePptxTextLayout={
 manifest:{version:1,manifestId:'geometry-only',revision:'1',faces:[{faceId:'unused',family:'Unused',weight:400,style:'normal',stretch:100,source:{kind:'host',resourceId:'unused',contentDigest:`sha256:${'0'.repeat(64)}`}}],fallbackChains:[]},
 resolver:{providerId:'unused',providerRevision:'1',resolve(){throw Error('No text')},load(){throw Error('No font')}},shaper:{providerId:'unused',providerRevision:'1',shape(){throw Error('No glyphs')}},defaults:{fontFamilies:['Unused'],fontSizeHundredthPt:1200,script:'Latn',language:'en-US',direction:'ltr'},
}
function fixture(){
 const deck=JSON.parse(readFileSync(resolve(import.meta.dirname,'../../../go/pptxpatch/testdata/native-contract/valid/parsed-full.json'),'utf8')) as NativePptxDeck
 const shape=deck.slides[0]!.elements.find(e=>e.kind==='shape')!
 if(shape.kind!=='shape')throw Error('missing shape')
 deck.slides[0]!.elements=[shape];delete shape.preset;shape.paragraphs=[];delete shape.textBody
 shape.compatibility={status:'preserveOnly',diagnostics:[{severity:'warning',code:'pptx.custom-geometry-preview',message:'Read-only source paths'}]};deck.slides[0]!.compatibility.status='preserveOnly';deck.compatibility.status='preserveOnly'
 shape.geometry={profile:'drawingml-paths-v1',textRect:{x:100,y:200,cx:800,cy:600},paths:[
 {fillMode:'norm',stroke:false,commands:[{kind:'moveTo',x:0,y:0},{kind:'quadBezierTo',x1:500,y1:1000,x:1000,y:0},{kind:'cubicBezierTo',x1:1000,y1:250,x2:500,y2:750,x:0,y:1000},{kind:'close'}]},
 {fillMode:'none',stroke:true,commands:[{kind:'moveTo',x:1000,y:0},{kind:'arcTo',rx:1000,ry:500,largeArc:false,clockwise:true,x:-1000,y:0},{kind:'arcTo',rx:1000,ry:500,largeArc:false,clockwise:true,x:1000,y:0},{kind:'close'}]},
 ]}
 shape.fill='336699';shape.stroke={color:'000000',widthEmu:10}
 return {deck,shape}
}
it('compiles public custom geometry and emits ordered independent path paint',async()=>{
 const {deck,shape}=fixture(),before=JSON.stringify(deck)
 expect(validateNativePptx(deck).ok,JSON.stringify(validateNativePptx(deck))).toBe(true)
 const tree=await compileNativePptxSlide(deck,0,{textLayout:layout})
 const surface=createRecordingPaintSurface();paintSlideRenderTree(tree,surface)
 const paths=surface.finish().filter(p=>p.kind==='path')
 expect(paths).toHaveLength(2)
 expect(paths[0]!.path).toEqual(shape.geometry!.paths[0]!.commands)
 expect(paths[0]!.fill).toBe('336699');expect(paths[0]!.stroke).toBeUndefined()
 expect(paths[1]!.fill).toBeUndefined();expect(paths[1]!.stroke?.widthEmu).toBe(10)
 expect(paths[1]!.path[1]).toMatchObject({kind:'arcTo',clockwise:true,rx:1000,ry:500})
 expect(JSON.stringify(deck)).toBe(before)
})
it.each([
 ['missing endpoint',(g:NativeEvaluatedGeometry)=>{delete g.paths[0]!.commands[0]!.x}],
 ['irrelevant field',(g:NativeEvaluatedGeometry)=>{g.paths[0]!.commands[0]!.rx=10}],
 ['wrong initial command',(g:NativeEvaluatedGeometry)=>{g.paths[0]!.commands[0]!.kind='lineTo'}],
 ['implicit radius correction',(g:NativeEvaluatedGeometry)=>{g.paths[1]!.commands[1]!.rx=1}],
 ['long arc',(g:NativeEvaluatedGeometry)=>{g.paths[1]!.commands[1]!.largeArc=true}],
 ['collapsed text rectangle',(g:NativeEvaluatedGeometry)=>{g.textRect.cx=0}],
 ['path cap',(g:NativeEvaluatedGeometry)=>{g.paths[0]!.commands.push(...Array(512).fill({kind:'close'}))}],
 ['total cap',(g:NativeEvaluatedGeometry)=>{g.paths=Array(17).fill({fillMode:'norm',stroke:true,commands:[{kind:'moveTo',x:0,y:0},...Array(511).fill({kind:'close'})]})}],
 ] as const)('refuses %s before paint',async(_name,mutate)=>{
 const {deck,shape}=fixture();mutate(shape.geometry!)
 expect(validateNativePptx(deck).ok).toBe(false)
 await expect(compileNativePptxSlide(deck,0,{textLayout:layout})).rejects.toThrow()
})
it('requires one geometry source and read-only authority for parsed and authored input',()=>{
 const {deck,shape}=fixture();shape.preset='rect';expect(validateNativePptx(deck).ok).toBe(false);delete shape.preset
 shape.compatibility.status='editable';expect(validateNativePptx(deck).ok).toBe(false)
 shape.provenance='authored';delete shape.source;shape.passthrough=[];expect(validateNativePptx(deck).ok).toBe(false)
})
it('honors caller coordinate bounds for controls outside the shape frame',async()=>{
 const {deck,shape}=fixture();shape.geometry!.paths[0]!.commands[1]!.x1=281474976710656
 expect(validateNativePptx(deck).ok,JSON.stringify(validateNativePptx(deck))).toBe(true)
 await expect(compileNativePptxSlide(deck,0,{textLayout:layout})).rejects.toThrow(/coordinate/i)
 shape.geometry!.paths[0]!.commands[1]!.x1=100000000
 await expect(compileNativePptxSlide(deck,0,{textLayout:layout,maxCoordinateEmu:99999999})).rejects.toThrow(/coordinate/i)
})

it('bounds a custom text rectangle in cumulative group space independently of path controls',async()=>{
 const {deck,shape}=fixture()
 shape.transform={x:0,y:0,cx:100,cy:100}
 shape.geometry!.textRect={x:1000000,y:1000000,cx:1000,cy:1000}
 const source=JSON.parse(readFileSync(resolve(import.meta.dirname,'../../../go/pptxpatch/testdata/native-contract/valid/parsed-full.json'),'utf8')) as NativePptxDeck
 const group=source.slides[0]!.elements.find(e=>e.kind==='group')!
 if(group.kind!=='group')throw Error('missing group')
 group.transform={x:0,y:0,cx:20000,cy:20000};group.childTransform={x:0,y:0,cx:100,cy:100};group.children=[shape]
 group.compatibility={status:'preserveOnly',diagnostics:[{severity:'warning',code:'fixture.group',message:'Read-only geometry child'}]}
 deck.slides[0]!.elements=[group]
 expect(validateNativePptx(deck).ok,JSON.stringify(validateNativePptx(deck))).toBe(true)
 await expect(compileNativePptxSlide(deck,0,{textLayout:layout,maxCoordinateEmu:100000000})).rejects.toThrow(/world-space bound/)
 shape.geometry!.textRect={x:0,y:0,cx:100,cy:100}
 await expect(compileNativePptxSlide(deck,0,{textLayout:layout,maxCoordinateEmu:100000000})).resolves.toMatchObject({nodes:[{kind:'group'}]})
})

it('paints all six source path fill modes with explicit tone-policy diagnostics',async()=>{
 const {deck,shape}=fixture()
 const modes=['norm','none','darken','darkenLess','lighten','lightenLess'] as const
 shape.fill='000000';shape.geometry!.paths=modes.map(fillMode=>({...shape.geometry!.paths[0]!,fillMode}))
 const before=JSON.stringify(deck)
 const tree=await compileNativePptxSlide(deck,0,{textLayout:layout})
 expect(tree.diagnostics.some(d=>d.code==='geometry.deterministicPathTone'&&d.message.includes('linear-srgb-path-tone-20-40-v1'))).toBe(true)
 const surface=createRecordingPaintSurface();paintSlideRenderTree(tree,surface)
 expect(surface.finish().filter(p=>p.kind==='path').map(p=>p.fill)).toEqual(['000000',undefined,'000000','000000','AAAAAA','7C7C7C'])
 expect(JSON.stringify(deck)).toBe(before)
 shape.geometry!.paths[0]!.fillMode='unknown' as never
 expect(validateNativePptx(deck).ok).toBe(false)
 await expect(compileNativePptxSlide(deck,0,{textLayout:layout})).rejects.toThrow()
})

it('retains callout paths outside the frame while clipping only legacy text',async()=>{
 const {deck,shape}=fixture();shape.transform={x:0,y:0,cx:1000,cy:1000}
 shape.geometry!.paths[0]!.commands[0]={kind:'moveTo',x:-500,y:-500}
 const tree=await compileNativePptxSlide(deck,0,{textLayout:layout})
 const node=tree.nodes[0]!;if(node.kind!=='shape')throw Error('missing shape')
 expect(node.clip).toBeUndefined()
 const modified={...tree,nodes:[{...node,textBody:{kind:'textBody' as const,sourceElementId:shape.id,bounds:node.bounds,fidelity:'legacyUnavailable' as const,status:'laidOut' as const,paragraphs:[]}}]}
 const surface=createRecordingPaintSurface();paintSlideRenderTree(modified,surface)
 const commands=surface.finish(),paths=commands.filter(c=>c.kind==='path')
 expect(paths[0]!.path[0]).toEqual({kind:'moveTo',x:-500,y:-500})
 const localClip=commands.findIndex(c=>c.kind==='clipRect'&&c.rect.cx===1000)
 expect(localClip).toBeGreaterThan(commands.lastIndexOf(paths[1]!))
})
