import {afterAll,describe,expect,it} from 'vitest'
import {readFileSync,writeFileSync,mkdtempSync,rmSync} from 'node:fs'
import {resolve} from 'node:path'
import {tmpdir} from 'node:os'
import {createHash} from 'node:crypto'
import {compilePptxPreview,previewStroke} from './compile.js'
import {decodePptxPreview} from './contract.js'
import {prepareNativeRasterResourceV1} from '@injoffice/docs/native-raster'
const root=resolve(import.meta.dirname,'../../..'),scratch=mkdtempSync(resolve(tmpdir(),'pptx-preview-worker-'))
afterAll(()=>rmSync(scratch,{recursive:true,force:true}))
const font=resolve(root,'node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf'),digest=`sha256:${createHash('sha256').update(readFileSync(font)).digest('hex')}`
const manifest=resolve(scratch,'fonts.json')
writeFileSync(manifest,JSON.stringify({version:1,faces:[{family:'DejaVu Sans',weight:400,style:'normal',path:font,sha256:digest}]}))
it('requires explicit source-frame opt-in and retains real glyph paint with a bounded approximation count',async()=>{
 const request=input(),element=request.deck.slides[0].elements[0]
 element.textBody.autoFit='shape-source-frame'
 element.compatibility={status:'preserveOnly',diagnostics:[{severity:'warning',code:'pptx.autofit-source-frame-approximate',message:'Source frame approximation'}]}
 await expect(compilePptxPreview(request)).rejects.toThrow('explicit preview opt-in')
 const before=JSON.stringify(request)
 const result=await compilePptxPreview({...request,source_frame_autofit_preview:true})
 expect(result.source_frame_autofit_count).toBe(1)
 expect(result.diagnostics.join(' ')).toContain('text.sourceFrameAutoFitApproximate')
 expect(JSON.stringify(result.nodes)).toContain('contentRun')
 expect(JSON.stringify(request)).toBe(before)
 for(const count of [-1,1.5,20001,'1',NaN])expect(()=>decodePptxPreview({...result,source_frame_autofit_count:count})).toThrow()
 await expect(compilePptxPreview({...request,source_frame_autofit_preview:'true'})).rejects.toThrow('boolean')
})
function input(){
 const deck=JSON.parse(readFileSync(resolve(root,'go/pptxpatch/testdata/native-contract/valid/parsed-full.json'),'utf8'))
 deck.assets=[];deck.slides=[deck.slides[0]];const element=deck.slides[0].elements[0];deck.slides[0].elements=[element]
 element.textBody={leftInsetEmu:0,rightInsetEmu:0,topInsetEmu:0,bottomInsetEmu:0,wrap:'none',verticalAnchor:'center',autoFit:'none',horizontalOverflow:'overflow',verticalOverflow:'overflow'}
 element.paragraphs=[{align:'left',level:0,bullet:false,runs:[{text:'Small ',fontFamily:'DejaVu Sans',fontSizeHundredthPt:1200,bold:false,italic:false,color:'123456'},{text:'large',fontFamily:'DejaVu Sans',fontSizeHundredthPt:2400,bold:false,italic:false,color:'123456'}]}]
 return {deck,slide_index:0,package_sha256:'a'.repeat(64),font_manifest_path:manifest}
}
describe('actual source-font native PPTX worker',()=>{
 it('keeps an element-scoped text refusal visible without hiding neighboring text',async()=>{
  const request=input(),slide=request.deck.slides[0],good=slide.elements[0],refused=structuredClone(good)
  refused.id='element:unsupported-text';refused.paragraphs=[];delete refused.textBody
  refused.compatibility={status:'refused',diagnostics:[{severity:'refusal',code:'pptx.text-content-unavailable',message:'Unsupported bullet font',scope:{slideId:slide.id,elementId:refused.id}}]}
  slide.elements=[refused,good];slide.compatibility={status:'refused',diagnostics:refused.compatibility.diagnostics};request.deck.compatibility=structuredClone(slide.compatibility)
  const result=await compilePptxPreview(request),paint=JSON.stringify(result.nodes)
  expect(paint).toContain('Unsupported text');expect(paint).toContain('contentRun');expect(paint).not.toContain('Slide rendering refused')
  expect(result.diagnostics.join(' ')).toContain('Unsupported bullet font')
 })
 it('keeps old boolean-only endpoints visibly unqualified while replaying typed sources',async()=>{
  const request=input(),source=JSON.parse(readFileSync(resolve(root,'go/pptxpatch/testdata/native-contract/valid/parsed-full.json'),'utf8')),connector=source.slides[0].elements.find((e:{kind:string})=>e.kind==='connector');request.deck.slides[0].elements=[connector];connector.tailArrow=true
  const legacy=await compilePptxPreview(request);expect(JSON.stringify(legacy.nodes)).toContain('Arrowhead unavailable');expect(JSON.stringify(legacy.nodes)).not.toContain('connectorArrow')
  connector.tailEnd={type:'diamond',w:'lg',len:'sm'};const typed=await compilePptxPreview(request);expect(JSON.stringify(typed.nodes)).toContain('connectorArrow');expect(typed.diagnostics.join(' ')).toContain('arrow.deterministicGeometry')
 })
 it('converts source miter thousandths-percent to SVG ratio without silently clamping',()=>{expect(previewStroke({color:'123456',widthEmu:12700,cap:'flat',join:'miter',miterLimit:800000})).toEqual({stroke:'123456',strokeWidth:12700,strokeLinecap:'butt',strokeLinejoin:'miter',strokeMiterlimit:8});expect(()=>previewStroke({color:'123456',widthEmu:12700,join:'miter',miterLimit:0})).toThrow('outside SVG replay range')})
 it('validates owned raster bytes and joins cropped nodes only to admitted resources',async()=>{
  const bytes=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=','base64'),before=Buffer.from(bytes)
  const resource=prepareNativeRasterResourceV1('ppt/media/image.png','image/png',bytes)
  expect(bytes).toEqual(before);expect(resource.width_px).toBe(1)
  const result=await compilePptxPreview(input()),image={kind:'image',rect:{x:0,y:0,cx:100,cy:100},resourceId:resource.id,crop:{left:50000,top:0,right:0,bottom:50000}}
  expect(()=>decodePptxPreview({...result,resources:[resource],nodes:[image]})).not.toThrow()
  expect(()=>decodePptxPreview({...result,resources:[],nodes:[image]})).toThrow()
  expect(()=>decodePptxPreview({...result,resources:[resource],nodes:[{...image,crop:{...image.crop,right:50000}}]})).toThrow()
  bytes[bytes.length-1]^=1;expect(()=>prepareNativeRasterResourceV1('ppt/media/image.png','image/png',bytes)).toThrow()
  expect(()=>decodePptxPreview({...result,resources:[{...resource,content_digest:`sha256:${'0'.repeat(64)}`}],nodes:[image]})).toThrow()
 })
 it('rejects string font weights before ambiguous face admission',async()=>{const bad=resolve(scratch,'weight.json');writeFileSync(bad,JSON.stringify({version:1,faces:[{family:'DejaVu Sans',weight:'400',style:'normal',path:font,sha256:digest}]}));await expect(compilePptxPreview({...input(),font_manifest_path:bad})).rejects.toThrow('Invalid operator face metadata')})
 it('compiles mixed metrics and anchored glyph outlines without mutating source',async()=>{const request=input(),before=JSON.stringify(request);const result=await compilePptxPreview(request);expect(result.font_digests).toEqual([digest]);expect(result.diagnostics.join(' ')).toContain('text.deterministicLayout');expect(JSON.stringify(result.nodes)).toContain('"kind":"path"');expect(JSON.stringify(request)).toBe(before);expect(decodePptxPreview(result)).toEqual(result)})
 it('refuses exact missing font faces instead of substitution',async()=>{const request=input();request.deck.slides[0].elements[0].paragraphs[0].runs[0].fontFamily='Missing Font';await expect(compilePptxPreview(request)).rejects.toThrow('Exact operator font unavailable: Missing Font')})
 it('rejects wrong configured font digests',async()=>{const bad=resolve(scratch,'bad.json');writeFileSync(bad,JSON.stringify({version:1,faces:[{family:'DejaVu Sans',weight:400,style:'normal',path:font,sha256:`sha256:${'0'.repeat(64)}`}]}));await expect(compilePptxPreview({...input(),font_manifest_path:bad})).rejects.toThrow('digest')})
 it('rejects hostile or unbounded vector paths before mounting',async()=>{const result=await compilePptxPreview(input());for(const d of ['M1e999 0','M0','<script>','M0 0LInfinity 1'])expect(()=>decodePptxPreview({...result,nodes:[{kind:'path',d,fill:'000000'}]})).toThrow()})
})
