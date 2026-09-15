import {afterAll,describe,expect,it} from 'vitest'
import {readFileSync,writeFileSync,mkdtempSync,rmSync} from 'node:fs'
import {resolve} from 'node:path'
import {tmpdir} from 'node:os'
import {createHash} from 'node:crypto'
import {compilePptxPreview,previewStroke} from './compile.js'
import {decodePptxPreview,type PreviewNode} from './contract.js'
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
it('requires explicit operator mapping and reports only source-bound painted substitutions',async()=>{
 const configured=resolve(scratch,'substitution-fonts.json')
 writeFileSync(configured,JSON.stringify({version:1,faces:[{family:'DejaVu Sans',weight:400,style:'normal',path:font,sha256:digest}],substitutions:{version:1,mappings:[{sourceFamily:'Missing Authored',targetFamily:'DejaVu Sans',weight:400,style:'normal'}]}}))
 const request=input();request.font_manifest_path=configured
 for(const r of request.deck.slides[0].elements[0].paragraphs[0].runs)r.fontFamily='Missing Authored'
 const before=JSON.stringify(request)
 await expect(compilePptxPreview(request)).rejects.toThrow('Exact operator font unavailable')
 for(const v of ['yes',1,null])await expect(compilePptxPreview({...request,font_substitution_preview:v})).rejects.toThrow('boolean')
 const painted=await compilePptxPreview({...request,font_substitution_preview:true})
 expect(painted.font_substitutions).toHaveLength(2)
 expect(painted.font_substitutions![0]).toMatchObject({source_id:request.deck.slides[0].elements[0].id,source_family:'Missing Authored',selected_family:'DejaVu Sans',font_digest:digest})
 expect(painted.font_substitution_policy).toBe('explicit-whole-run-font-substitution-v1')
 expect(painted.diagnostics.join(' ')).toContain('fontSubstitutionApproximate')
 expect(JSON.stringify(painted.nodes)).toContain('contentRun')
 expect(JSON.stringify(request)).toBe(before)
 for(const mutation of [{font_substitution_policy:undefined},{font_substitution_policy_sha256:'bad'},{font_substitutions:[]},{font_substitutions:[{...painted.font_substitutions![0],run_index:-1}]}])expect(()=>decodePptxPreview({...painted,...mutation})).toThrow()
 expect(()=>decodePptxPreview({...painted,font_digests:[]})).toThrow()
 expect(()=>decodePptxPreview({...painted,font_substitutions:[painted.font_substitutions![0],{...painted.font_substitutions![0],selected_family:'Forged'}]})).toThrow()
 const p=request.deck.slides[0].elements[0].paragraphs[0];Object.assign(p,{bullet:true,bulletCharacter:'q',bulletFontFamily:'Missing Authored'})
 const bullet=await compilePptxPreview({...request,font_substitution_preview:true})
 expect(JSON.stringify(bullet.nodes)).not.toContain('paragraphBullet')
 expect(bullet.font_substitutions).toBeUndefined()
})
it('binds inherited approximation to opt-in, read-only source and closed transport policy',async()=>{
 for(const value of ['yes',1,null])await expect(compilePptxPreview({...input(),inherited_text_preview:value})).rejects.toThrow('boolean')
 const request=input(),element=request.deck.slides[0].elements[0]
 element.compatibility={status:'preserveOnly',diagnostics:[{severity:'warning',code:'pptx.source-inherited-text-approximate',message:'source-latin-inheritance-approximate-v1; kerning disabled'}]}
 await expect(compilePptxPreview(request)).rejects.toThrow('explicit preview opt-in')
 const result=await compilePptxPreview({...request,inherited_text_preview:true})
 expect(result.inherited_text_preview_count).toBe(1);expect(result.inherited_text_policy).toBe('source-latin-inheritance-approximate-v1');expect(JSON.stringify(result.nodes)).toContain('contentRun')
 for(const count of [0,-1,1.5,20001,'1',NaN])expect(()=>decodePptxPreview({...result,inherited_text_preview_count:count})).toThrow()
 for(const policy of [undefined,'unknown'])expect(()=>decodePptxPreview({...result,inherited_text_policy:policy})).toThrow()
 expect(()=>decodePptxPreview({...result,inherited_text_preview_count:undefined})).toThrow()
 element.compatibility.status='editable';await expect(compilePptxPreview({...request,inherited_text_preview:true})).rejects.toThrow()
})
describe('actual source-font native PPTX worker',()=>{
 it('replays source crop within the rounded local clip and bounds untrusted radii',async()=>{
  const request=input(),source=JSON.parse(readFileSync(resolve(root,'go/pptxpatch/testdata/native-contract/valid/parsed-full.json'),'utf8'))
  const picture=source.slides[0].elements.find((e:{kind:string})=>e.kind==='picture')
  const bytes=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=','base64')
  const assetHash=createHash('sha256').update(bytes).digest('hex')
  request.deck.assets=[{id:picture.assetId,provenance:'parsed',contentType:'image/png',sha256:assetHash,byteLength:bytes.length,dataBase64:bytes.toString('base64'),source:{partName:'ppt/media/image.png',objectId:'asset',fingerprintSha256:assetHash},passthrough:[]}]
  picture.clip='roundRect';picture.crop={left:10000,top:20000,right:30000,bottom:0};picture.transform.cx=3000000;picture.transform.cy=2000000
  request.deck.slides[0].elements=[picture]
  const result=await compilePptxPreview(request),serialized=JSON.stringify(result.nodes)
  expect(serialized).toContain('"radius":333340');expect(serialized).toContain('"left":10000');expect(result.resources).toHaveLength(1)
  for(const radius of [-1,Infinity,1000001,'1'])expect(()=>decodePptxPreview({...result,nodes:[{kind:'group',transform:[1,0,0,1,0,0],clip:{x:0,y:0,cx:3000000,cy:2000000,radius},children:[]}]})).toThrow()
 })
 it('shapes authored markers in the exact supplied family and rejects a missing marker face',async()=>{
  const request=input(),p=request.deck.slides[0].elements[0].paragraphs[0]
  Object.assign(p,{bullet:true,bulletCharacter:'q',bulletFontFamily:'DejaVu Sans',marginLeftEmu:300000,indentEmu:-100000})
  const before=JSON.stringify(request)
  const rendered=await compilePptxPreview(request)
  expect(JSON.stringify(rendered.nodes)).toContain('paragraphBullet')
  expect(JSON.stringify(rendered.nodes)).toContain('contentRun')
  expect(JSON.stringify(request)).toBe(before)
  p.bulletFontEncoding='windows-symbol-byte-v1'
  const unqualified=await compilePptxPreview(request)
  expect(unqualified.diagnostics.join(' ')).toContain('symbol bullet font encoding is not qualified')
  expect(JSON.stringify(unqualified.nodes)).not.toContain('paragraphBullet')
  delete p.bulletFontEncoding
  p.bulletFontFamily='Missing Symbol Font'
  await expect(compilePptxPreview(request)).rejects.toThrow('Exact operator bullet font unavailable')
 })
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
 connector.geometry={profile:'drawingml-paths-v1',textRect:{x:0,y:0,cx:connector.transform.cx,cy:connector.transform.cy},paths:[{fillMode:'none',stroke:true,commands:[{kind:'moveTo',x:0,y:0},{kind:'lineTo',x:connector.transform.cx,y:0},{kind:'lineTo',x:connector.transform.cx,y:connector.transform.cy}]}]}
 connector.compatibility={status:'preserveOnly',diagnostics:[{severity:'warning',code:'pptx.connector-preset-preview',message:'catalog preview'}]};connector.transform={...connector.transform,rotationAngle:5400000};delete connector.flipH
 const bent=await compilePptxPreview(request),bentText=JSON.stringify(bent.nodes)
 expect(bentText).toContain('connectorArrow');expect(bentText).toMatch(/"d":"M0 0 L\d+ 0 L\d+ \d+(\.\d+)?"/);expect(bent.diagnostics.join(' ')).toContain('connector.presetShaftPreview')
 // A short terminal segment (cy below the arrow-v1 inset) paints untrimmed under the arrowhead; the slide never fails.
 connector.geometry.paths[0].commands[2]={kind:'lineTo',x:connector.transform.cx,y:connector.stroke.widthEmu}
 const short=await compilePptxPreview(request)
 expect(JSON.stringify(short.nodes)).toContain('connectorArrow');expect(short.diagnostics.join(' ')).toContain('connector.shaftInsetSkipped');expect(JSON.stringify(short.nodes)).not.toContain('Connector arrow unavailable')
 // A fully degenerate shaft routes to the element placeholder instead of rejecting the slide.
 connector.geometry.paths[0].commands=[{kind:'moveTo',x:0,y:0},{kind:'lineTo',x:0,y:0}]
 const degenerate=await compilePptxPreview(request)
 expect(JSON.stringify(degenerate.nodes)).toContain('Connector arrow unavailable');expect(JSON.stringify(degenerate.nodes)).not.toContain('connectorArrow');expect(degenerate.diagnostics.join(' ')).toContain('connector.arrowUnavailable')
 connector.geometry.paths[0].commands=[{kind:'moveTo',x:0,y:0},{kind:'lineTo',x:connector.transform.cx,y:0},{kind:'close'}]
 await expect(compilePptxPreview(request)).rejects.toThrow()
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

it('paints centered and bottom table labels with real supplied font outlines and declared line policy', async () => {
 const results = []
 for (const verticalAnchor of ['top', 'center', 'bottom']) {
  const request = input(), source = request.deck.slides[0].elements[0]
  request.deck.slides[0].elements = [{id:source.id,provenance:source.provenance,source:source.source,passthrough:source.passthrough,compatibility:source.compatibility,kind:'table',
   transform:{x:0,y:0,cx:3000000,cy:3000001}, table:{columnWidths:[3000000],rowHeights:[3000001],rows:[[{text:'Anchor',fill:'FFFFFF',
    paragraphs:[{align:'left',level:0,bullet:false,runs:[{text:'Anchor',fontFamily:'DejaVu Sans',fontSizeHundredthPt:1200,color:'123456'}]}],
    textBody:{...source.textBody,verticalAnchor,leftInsetEmu:10000,rightInsetEmu:20000,topInsetEmu:30000,bottomInsetEmu:40000},
   }]]}}]
  const before = JSON.stringify(request), result = await compilePptxPreview(request)
  expect(JSON.stringify(request)).toBe(before)
  expect(result.policy).toBe('max-run-natural-v1')
  expect(result.diagnostics.join(' ')).toContain('text.deterministicLayout')
  expect(JSON.stringify(result.nodes)).not.toContain('placeholder')
  const runs: Extract<PreviewNode,{kind:'group'}>[] = []
  const visit = (nodes:PreviewNode[]) => {for(const node of nodes) if(node.kind==='group'){if(node.sourceRole==='contentRun')runs.push(node);visit(node.children)}}
  visit(result.nodes)
  expect(runs).toHaveLength(6)
  expect(runs[0]!.children.some(n=>n.kind==='path')).toBe(true)
  results.push(runs[0]!.transform[5])
 }
 expect(results[1]).toBeGreaterThan(results[0]!)
 expect(results[2]).toBeGreaterThan(results[1]!)
 expect(Math.abs((results[2]!-results[0]!)-2*(results[1]!-results[0]!))).toBeLessThanOrEqual(1)
})
it('still emits a slide raster when remaining content is an explicit source refusal',async()=>{
 const request=input()
 request.deck.slides[0].elements=[]
 request.deck.slides[0].compatibility={status:'preserveOnly',diagnostics:[{severity:'warning',code:'pptx.unsupported-shape',message:'placeholder inheritance remains unqualified'}]}
 request.deck.slides[0].passthrough=[{token:'pass-unqualified',ownerPart:'ppt/slides/slide1.xml',fingerprintSha256:'4'.repeat(64),disposition:'preserve'}]
 const result=await compilePptxPreview(request)
 expect(result.slide_count).toBe(1)
 expect(result.width).toBe(12192000)
 expect(result.height).toBe(6858000)
 expect(result.background).toBe('FFFFFF')
 expect(result.diagnostics.join(' ')).toContain('placeholder inheritance remains unqualified')
 expect(result.diagnostics.join(' ')).toContain('render.preserveOnly')
})
