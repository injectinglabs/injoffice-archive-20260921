import {afterAll,expect,it} from 'vitest'
import {readFileSync,writeFileSync,mkdtempSync,rmSync} from 'node:fs'
import {resolve} from 'node:path'
import {tmpdir} from 'node:os'
import {createHash} from 'node:crypto'
import {compilePptxPreview} from './compile.js'
import {decodePptxPreview,type PreviewNode} from './contract.js'
const root=resolve(import.meta.dirname,'../../..'),scratch=mkdtempSync(resolve(tmpdir(),'pptx-custom-geometry-'))
afterAll(()=>rmSync(scratch,{recursive:true,force:true}))
const font=resolve(root,'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf'),digest=`sha256:${createHash('sha256').update(readFileSync(font)).digest('hex')}`
const manifest=resolve(scratch,'fonts.json')
writeFileSync(manifest,JSON.stringify({version:1,faces:[{family:'DejaVu Sans',weight:400,style:'normal',path:font,sha256:digest}]}))
function input(){
 const deck=JSON.parse(readFileSync(resolve(root,'go/pptxpatch/testdata/native-contract/valid/parsed-full.json'),'utf8'))
 const shape=deck.slides[0].elements.find((e:{kind:string})=>e.kind==='shape');deck.slides[0].elements=[shape];deck.assets=[]
 delete shape.preset;shape.transform={x:0,y:0,cx:4000000,cy:3000000}
 shape.compatibility={status:'preserveOnly',diagnostics:[{severity:'warning',code:'pptx.custom-geometry-preview',message:'Read-only paths'}]}
 shape.fill='336699';shape.stroke={color:'000000',widthEmu:10000}
 shape.geometry={profile:'drawingml-paths-v1',textRect:{x:1000000,y:1000000,cx:2000000,cy:1000000},paths:[{fillMode:'norm',stroke:true,commands:[
 {kind:'moveTo',x:4000000,y:1500000},{kind:'arcTo',rx:2000000,ry:1500000,largeArc:false,clockwise:true,x:0,y:1500000},{kind:'arcTo',rx:2000000,ry:1500000,largeArc:false,clockwise:true,x:4000000,y:1500000},{kind:'close'},
 {kind:'moveTo',x:2500000,y:1500000},{kind:'arcTo',rx:500000,ry:250000,largeArc:false,clockwise:false,x:1500000,y:1500000},{kind:'arcTo',rx:500000,ry:250000,largeArc:false,clockwise:false,x:2500000,y:1500000},{kind:'close'}
 ]},{fillMode:'none',stroke:true,commands:[{kind:'moveTo',x:0,y:0},{kind:'quadBezierTo',x1:2000000,y1:3000000,x:4000000,y:0},{kind:'cubicBezierTo',x1:4000000,y1:1000000,x2:1000000,y2:2000000,x:0,y:3000000}]}]}
 shape.textBody={leftInsetEmu:100000,rightInsetEmu:100000,topInsetEmu:100000,bottomInsetEmu:100000,wrap:'none',verticalAnchor:'top',autoFit:'none',horizontalOverflow:'overflow',verticalOverflow:'overflow'}
 shape.paragraphs=[{align:'left',level:0,bullet:false,runs:[{text:'A',fontFamily:'DejaVu Sans',fontSizeHundredthPt:1200,bold:false,italic:false,color:'123456'}]}]
 return {deck,slide_index:0,package_sha256:'a'.repeat(64),font_manifest_path:manifest}
}
const flatten=(nodes:PreviewNode[]):PreviewNode[]=>nodes.flatMap(n=>[n,...(n.kind==='group'?flatten(n.children):[])])
it('replays arcs with a winding hole, both Bezier kinds, and real glyphs in evaluated text rect',async()=>{
 const request=input(),before=JSON.stringify(request),paint=await compilePptxPreview(request),nodes=flatten(paint.nodes)
 const paths=nodes.filter(n=>n.kind==='path')
 expect(paths.some(p=>p.d.includes('A2000000 1500000 0 0 1'))).toBe(true)
 expect(paths.some(p=>p.d.includes('A500000 250000 0 0 0'))).toBe(true)
 expect(paths.some(p=>p.d.includes('Q2000000 3000000 4000000 0 C4000000 1000000 1000000 2000000 0 3000000'))).toBe(true)
 const glyph=nodes.find(n=>n.kind==='group'&&n.sourceRole==='contentRun')
 expect(glyph?.kind).toBe('group');if(glyph?.kind!=='group')throw Error('missing actual glyph')
 expect(glyph.transform[4]).toBe(1100000);expect(glyph.transform[5]).toBeGreaterThan(1100000);expect(glyph.transform[5]).toBeLessThan(2000000)
 expect(paint.font_digests).toContain(digest)
 expect(JSON.stringify(request)).toBe(before)
 for(const d of ['M0 0 A1 1 0 0 1 3 0','M0 0 A1 1 45 0 1 1 1','M0 0 A1 1 0 2 1 1 1','M0 0 A0 1 0 0 1 1 1','A1 1 0 0 1 1 1','M0 0 A1 1 0 0 1 1.5 1']){
  expect(()=>decodePptxPreview({...paint,nodes:[{kind:'path',d,fill:'336699'}]})).toThrow()
 }
})
it('keeps legacy paragraphs on the old full-frame text origin when textBody is absent',async()=>{
 const request=input();delete request.deck.slides[0].elements[0].textBody
 const nodes=flatten((await compilePptxPreview(request)).nodes),glyph=nodes.find(n=>n.kind==='group'&&n.sourceRole==='contentRun')
 if(glyph?.kind!=='group')throw Error('missing glyph');expect(glyph.transform[4]).toBe(0)
})
