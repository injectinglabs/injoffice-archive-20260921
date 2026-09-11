import {readFileSync,statSync} from 'node:fs'
import {isAbsolute} from 'node:path'
import {createHash} from 'node:crypto'
import {assertNativePptx,type NativePptxDeck,type NativeElement,type NativeParagraph} from '@injoffice/pptx-native'
import {compileNativePptxSlide,createRecordingPaintSurface,paintSlideRenderTree,type RenderPathCommand,type RenderStroke} from '@injoffice/pptx-render'
import {createHarfBuzzTextShaperV1,createHarfBuzzOutlineProviderV1,inspectHarfBuzzFontMetricsV1} from '@injoffice/font-metrics/harfbuzz'
import type {NativeFontManifest,NativeFontResolver,ResolvedFontFace,FontResource} from '@injoffice/font-metrics/layout'
import {decodePptxPreview,type PreviewNode,type PptxPreview,type PreviewStroke} from './contract.js'
import {prepareNativeRasterResourceV1,type NativeDocxPagePaintMediaAssetV1} from '@injoffice/docs/native-raster'
import {previewArrow,previewArrowShaftInset} from './arrows.js'

const hash=(bytes:Uint8Array)=>`sha256:${createHash('sha256').update(bytes).digest('hex')}` as const
const object=(v:unknown):Record<string,unknown>=>{if(!v||typeof v!=='object'||Array.isArray(v))throw new TypeError('Expected bounded object');return v as Record<string,unknown>}
export function previewStroke(stroke:RenderStroke):PreviewStroke {
 const ratio=stroke.miterLimit===undefined?undefined:stroke.miterLimit/100000
 if(ratio!==undefined&&(!Number.isFinite(ratio)||ratio<1))throw new Error('DrawingML miter limit is outside SVG replay range')
 return {stroke:stroke.color,strokeWidth:stroke.widthEmu,strokeLinecap:stroke.cap==='flat'?'butt':stroke.cap,strokeLinejoin:stroke.join,strokeMiterlimit:ratio}
}
function fontProviders(path:string){
 if(!isAbsolute(path)||statSync(path).size>65536)throw new Error('Operator font manifest must be an absolute bounded local file')
 const config=object(JSON.parse(readFileSync(path,'utf8')))
 if(config.version!==1||!Array.isArray(config.faces)||config.faces.length>32)throw new Error('Invalid operator font manifest')
 const resources=new Map<string,FontResource>(),outlines=new Map<string,ReturnType<typeof createHarfBuzzOutlineProviderV1>>()
 const faces:NativeFontManifest['faces'][number][]=[]
 const manifest:NativeFontManifest={version:1,manifestId:'pptx-preview-fonts',revision:hash(Buffer.from(JSON.stringify(config))),faces,fallbackChains:[]}
 let total=0
 for(const [index,entry] of config.faces.entries()){
  const f=object(entry)
  if(typeof f.family!=='string'||!f.family||f.family.length>128||typeof f.weight!=='number'||![400,700].includes(f.weight)||typeof f.style!=='string'||!['normal','italic'].includes(f.style)||typeof f.path!=='string'||!isAbsolute(f.path)||typeof f.sha256!=='string'||!/^sha256:[a-f0-9]{64}$/.test(f.sha256))throw new Error('Invalid operator face metadata')
  if(statSync(f.path).size>16*1024*1024)throw new Error('Operator font exceeds 16 MiB')
  const bytes=Uint8Array.from(readFileSync(f.path));total+=bytes.length
  if(total>64*1024*1024||hash(bytes)!==f.sha256)throw new Error('Operator font digest or total byte budget failed')
  const id=`font-${index}`, digest=f.sha256 as `sha256:${string}`
  if(manifest.faces.some(face=>face.family===f.family&&face.weight===f.weight&&face.style===f.style))throw new Error('Ambiguous operator face identity')
  const face:ResolvedFontFace={faceId:id,family:f.family,weight:Number(f.weight),style:f.style as 'normal'|'italic',stretch:100,sourceKind:'host',resourceId:id,contentDigest:digest,resolution:'exact',matchedFamily:f.family}
  faces.push({faceId:id,family:face.family,weight:face.weight,style:face.style,stretch:100,source:{kind:'host',resourceId:id,contentDigest:digest}})
  const request={bytes,contentDigest:digest}
  resources.set(id,{face,bytes,metrics:inspectHarfBuzzFontMetricsV1(request)})
  outlines.set(id,createHarfBuzzOutlineProviderV1(request))
 }
 const resolver:NativeFontResolver={providerId:'injoffice.pptx.operator-fonts',providerRevision:'v1',resolve({run}){
  const found=[...resources.values()].find(resource=>run.font.families.includes(resource.face.family)&&run.font.weight===resource.face.weight&&run.font.style===resource.face.style)
  return found?{status:'resolved',face:found.face,attemptedFaceIds:[found.face.faceId],decisions:[]}:{status:'refused',attemptedFaceIds:[],decisions:[{code:'font-not-found',message:`Exact operator font unavailable: ${run.font.families.join(', ')} / ${run.font.weight} / ${run.font.style}`,recoverable:false}]}
 },load(face){const resource=resources.get(face.faceId);if(!resource)throw new Error('Unresolved font resource');return resource}}
 return {manifest,resolver,outlines,resources}
}

export async function compilePptxPreview(input:unknown):Promise<PptxPreview>{
 const request=object(input)
 if(typeof request.package_sha256!=='string'||!/^[a-f0-9]{64}$/.test(request.package_sha256)||typeof request.font_manifest_path!=='string'||!Number.isSafeInteger(request.slide_index))throw new Error('Invalid source-bound preview input')
 assertNativePptx(request.deck)
 const deck=request.deck as NativePptxDeck
 if(deck.origin!=='parsed'||request.slide_index as number<0||request.slide_index as number>=deck.slides.length)throw new Error('Preview requires a parsed source slide')
 const fonts=fontProviders(request.font_manifest_path)
 const checkParagraphs=(paragraphs:readonly NativeParagraph[])=>{for(const paragraph of paragraphs)for(const run of paragraph.runs){if(![...fonts.resources.values()].some(r=>r.face.family===run.fontFamily&&r.face.weight===(run.bold?700:400)&&r.face.style===(run.italic?'italic':'normal')))throw new Error(`Exact operator font unavailable: ${run.fontFamily??'unresolved family'} / ${run.bold?'bold':'regular'} / ${run.italic?'italic':'normal'}`)}}
 const checkElements=(elements:readonly NativeElement[])=>{for(const element of elements){if(element.kind==='text'||element.kind==='shape')checkParagraphs(element.paragraphs);if(element.kind==='group')checkElements(element.children);if(element.kind==='table')for(const row of element.table.rows)for(const cell of row)if(cell.paragraphs)checkParagraphs(cell.paragraphs)}}
 checkElements(deck.slides[request.slide_index as number]!.elements)
 const tree=await compileNativePptxSlide(deck,request.slide_index as number,{lineLayoutPolicy:'max-run-natural-v1',maxGlyphs:20000,maxNodes:20000,textLayout:{manifest:fonts.manifest,resolver:fonts.resolver,shaper:createHarfBuzzTextShaperV1({sourceRevision:'pptx-preview-v1'}),defaults:{fontFamilies:[],fontSizeHundredthPt:1200,script:'Latn',language:'en-US',direction:'ltr'}}})
 const recording=createRecordingPaintSurface();paintSlideRenderTree(tree,recording)
 const root:Extract<PreviewNode,{kind:'group'}>={kind:'group',transform:[1,0,0,1,0,0],children:[]}
 const stack=[root],diagnostics=tree.diagnostics.map(d=>`${d.code}: ${d.message}`)
 let glyphs=0
 const resources=new Map<string,NativeDocxPagePaintMediaAssetV1>()
 for(const command of recording.finish()){
  const current=stack[stack.length-1]!
  switch(command.kind){
   case 'save':{const group:typeof root={kind:'group',transform:[1,0,0,1,0,0],children:[]};current.children.push(group);stack.push(group);break}
   case 'restore':if(stack.length<2)throw new Error('Paint stack underflow');stack.pop();break
   case 'transform':{const t=command.transform;current.transform=[t.aPpm/1e6,t.bPpm/1e6,t.cPpm/1e6,t.dPpm/1e6,t.txEmu,t.tyEmu];break}
   case 'clipRect':current.clip=command.rect;break
   case 'path':{
    if(command.headArrow&&!command.headEnd||command.tailArrow&&!command.tailEnd){diagnostics.push('Boolean-only arrowheads lack source type/dimensions and remain unqualified');current.children.push({kind:'placeholder',rect:{x:0,y:0,cx:300000,cy:100000},label:'Arrowhead unavailable'});break}
    const stroke=command.stroke
    const paint={fill:command.fill??'none',...(stroke?previewStroke(stroke):{})}
    const first=command.path[0]
    if(first?.kind==='rect'||first?.kind==='roundRect')current.children.push({kind:'rect',rect:first.rect,radius:first.kind==='roundRect'?first.radiusEmu:0,...paint})
    else if(first?.kind==='ellipse')current.children.push({kind:'ellipse',rect:first.rect,...paint})
    else {
     let d=command.path.map(pathPart).join(' ')
     if((command.headEnd||command.tailEnd)&&stroke){const start=command.path[0],finish=command.path[1];if(command.path.length!==2||start?.kind!=='moveTo'||finish?.kind!=='lineTo')throw new Error('Arrow shaft must be a straight source connector');const length=Math.hypot(finish.x-start.x,finish.y-start.y),head=previewArrowShaftInset(command.headEnd,stroke.widthEmu),tail=previewArrowShaftInset(command.tailEnd,stroke.widthEmu);if(length<=head+tail)throw new Error('Source connector is too short for arrow-v1 endpoint geometry');const x=(finish.x-start.x)/length,y=(finish.y-start.y)/length;d=`M${start.x+x*head} ${start.y+y*head} L${finish.x-x*tail} ${finish.y-y*tail}`}
     current.children.push({kind:'path',d,...paint})
    }
    if(command.headEnd?.type!=='none'&&command.headEnd||command.tailEnd?.type!=='none'&&command.tailEnd){
     const start=command.path[0],finish=command.path[1]
     if(command.path.length!==2||start?.kind!=='moveTo'||finish?.kind!=='lineTo'||!stroke)throw new Error('Typed arrows require a source-bound straight stroked connector')
     for(const [end,tip,other] of [[command.headEnd,start,finish],[command.tailEnd,finish,start]] as const){if(!end)continue;const arrow=previewArrow(end,tip,{x:tip.x-other.x,y:tip.y-other.y},stroke.widthEmu,stroke.color);if(arrow)current.children.push(arrow)}
     diagnostics.push('arrow.deterministicGeometry: InjOffice arrow-v1 uses source type and named widths/lengths (2/3/5 × stroke; omitted = medium), not Office-equivalent geometry')
    }
    break
   }
   case 'glyphRun':{
    const run=command.run,provider=fonts.outlines.get(run.faceId??'')
    if(!provider||fonts.resources.get(run.faceId!)?.face.contentDigest!==run.contentDigest)throw new Error('Glyph outline font identity mismatch')
    for(const glyph of run.glyphs){
     if(++glyphs>20000)throw new Error('Glyph path budget exceeded')
     const outline=provider.outline(glyph.glyphId),scale=run.fontSizeMilliPoints*12.7/outline.units_per_em
     const d=outline.path.map(p=>{switch(p.kind){case 'move_to':return `M${p.x} ${p.y}`;case 'line_to':return `L${p.x} ${p.y}`;case 'quadratic_to':return `Q${p.control_x} ${p.control_y} ${p.x} ${p.y}`;case 'cubic_to':return `C${p.control_1_x} ${p.control_1_y} ${p.control_2_x} ${p.control_2_y} ${p.x} ${p.y}`;case 'close_path':return 'Z'}}).join(' ')
     current.children.push({kind:'group',sourceRole:run.sourceRole??'contentRun',transform:[scale,0,0,-scale,run.x+glyph.xEmu,run.baselineY+glyph.yEmu],children:[{kind:'path',d,fill:run.color}]})
    }
    break
   }
   case 'placeholder':current.children.push({kind:'placeholder',rect:command.rect,label:command.label});break
   case 'image':{
    let resource=resources.get(command.assetId)
    if(!resource){
     const asset=deck.assets.find(a=>a.id===command.assetId)
     if(!asset?.source?.partName||!asset.dataBase64||!['image/png','image/jpeg'].includes(asset.contentType)||asset.dataBase64.length>24*1024*1024)throw new Error('Native image requires bounded embedded PNG or baseline JFIF bytes')
     const bytes=Buffer.from(asset.dataBase64,'base64')
     if(bytes.toString('base64')!==asset.dataBase64||bytes.length!==asset.byteLength||hash(bytes).slice(7)!==asset.sha256.replace(/^sha256:/,''))throw new Error('Native image source digest or byte identity mismatch')
     resource=prepareNativeRasterResourceV1(asset.source.partName,asset.contentType as 'image/png'|'image/jpeg',bytes)
     resources.set(command.assetId,resource)
    }
    current.children.push({kind:'image',rect:command.rect,resourceId:resource.id,...(command.crop?{crop:{...command.crop}}:{})});break
   }
  }
 }
 const result:PptxPreview={version:1,package_sha256:request.package_sha256,slide_index:request.slide_index as number,slide_count:deck.slides.length,width:tree.size.cx,height:tree.size.cy,background:tree.background.color,policy:'max-run-natural-v1',nodes:root.children,diagnostics,font_digests:[...fonts.resources.values()].map(r=>r.face.contentDigest),resources:[]}
 result.resources=[...resources.values()].sort((a,b)=>a.part_name.toLowerCase()<b.part_name.toLowerCase()?-1:1)
 return decodePptxPreview(result)
}
function pathPart(p:RenderPathCommand):string {switch(p.kind){case 'moveTo':return `M${p.x} ${p.y}`;case 'lineTo':return `L${p.x} ${p.y}`;case 'close':return 'Z';default:throw new Error('Unmodeled path command')}}
