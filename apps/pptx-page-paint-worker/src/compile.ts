import {workbookChartsForPreview} from './workbookCharts.js'
import {chartGlyphExtents} from './chartGlyphExtents.js'
import {readFileSync,statSync} from 'node:fs'
import {isAbsolute} from 'node:path'
import {createHash} from 'node:crypto'
import {assertNativePptx,type NativePptxDeck,type NativeElement,type NativeParagraph} from '@injoffice/pptx-native'
import {renderTransformMatrix,SourceAffineBudget,compileNativePptxSlide,createRecordingPaintSurface,paintSlideRenderTree,type RenderPathCommand,type RenderStroke} from '@injoffice/pptx-render'
import {createHarfBuzzTextShaperV1,createHarfBuzzOutlineProviderV1,inspectHarfBuzzFontMetricsV1} from '@injoffice/font-metrics/harfbuzz'
import type {NativeFontManifest,NativeFontResolver,ResolvedFontFace,FontResource} from '@injoffice/font-metrics/layout'
import {decodeExplicitFontPolicyV1,selectExplicitFontV1,EXPLICIT_FONT_POLICY_V1} from '@injoffice/font-metrics/layout'
import {decodePptxPreview,type PreviewNode,type PptxPreview,type PreviewStroke} from './contract.js'
import {prepareNativeRasterResourceV1,type NativeDocxPagePaintMediaAssetV1} from '@injoffice/docs/native-raster'
import {previewArrow,previewArrowStrokeBase} from './arrows.js'
import {previewConnectorShaft} from './connectorShaft.js'

const previewColor=(color:string)=>/^#[0-9A-F]{6}$/.test(color)?color.slice(1):color
const hash=(bytes:Uint8Array)=>`sha256:${createHash('sha256').update(bytes).digest('hex')}` as const
const object=(v:unknown):Record<string,unknown>=>{if(!v||typeof v!=='object'||Array.isArray(v))throw new TypeError('Expected bounded object');return v as Record<string,unknown>}
export function previewStroke(stroke:RenderStroke):PreviewStroke {
 const ratio=stroke.miterLimit===undefined?undefined:stroke.miterLimit/100000
 if(ratio!==undefined&&(!Number.isFinite(ratio)||ratio<1))throw new Error('DrawingML miter limit is outside SVG replay range')
 return {stroke:previewColor(stroke.color),strokeWidth:stroke.widthEmu,strokeLinecap:stroke.cap==='flat'?'butt':stroke.cap,strokeLinejoin:stroke.join,strokeMiterlimit:ratio}
}
// The operator font manifest's face cap. The Go helper re-encodes this same
// number in validatePPTXFontSubstitutions (pptx_font_substitutions.go); raising
// one without the other makes the worker accept a manifest the helper then
// refuses. 64 faces still sit far inside the surviving budgets: the manifest
// file is capped at 64 KiB (~230 faces at the ~280 bytes a face costs), the
// 64 MiB cumulative font-byte budget is unchanged, and the preview contract
// bounds font_digests — one per loaded face — at 256.
export const MAX_OPERATOR_FONT_FACES=64

function fontProviders(path:string,allowSubstitution=false){
 if(!isAbsolute(path)||statSync(path).size>65536)throw new Error('Operator font manifest must be an absolute bounded local file')
 const config=object(JSON.parse(readFileSync(path,'utf8')))
 // One generic message for four distinct conditions makes an operator
 // manifest that is merely too large indistinguishable from a corrupt one,
 // and the slide only ever shows the 422 body. Name the limit that failed.
 if(config.version!==1)throw new Error(`Invalid operator font manifest: version must be 1, got ${JSON.stringify(config.version)}`)
 if(!Array.isArray(config.faces))throw new Error('Invalid operator font manifest: faces must be an array')
 if(config.faces.length>MAX_OPERATOR_FONT_FACES)throw new Error(`Invalid operator font manifest: ${config.faces.length} faces exceeds the ${MAX_OPERATOR_FONT_FACES}-face limit`)
 const policy=config.substitutions===undefined?undefined:decodeExplicitFontPolicyV1(config.substitutions)
 const resources=new Map<string,FontResource>(),outlines=new Map<string,ReturnType<typeof createHarfBuzzOutlineProviderV1>>()
 const faces:NativeFontManifest['faces'][number][]=[]
 // Faces HarfBuzz cannot inspect are skipped rather than failing the slide
 // (a manifest may carry a face this deck never uses). Their families are
 // kept so a later "font unavailable" can say WHY the family is missing
 // instead of looking like the manifest never listed it — TrueType
 // collections (.ttc) land here.
 const skipped:string[]=[]
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
  const request={bytes,contentDigest:digest}
  let metrics
  try {
   metrics=inspectHarfBuzzFontMetricsV1(request)
  } catch {
   skipped.push(`${f.family} (${f.path})`)
   continue
  }
  faces.push({faceId:id,family:face.family,weight:face.weight,style:face.style,stretch:100,source:{kind:'host',resourceId:id,contentDigest:digest}})
  resources.set(id,{face,bytes,metrics})
  outlines.set(id,createHarfBuzzOutlineProviderV1(request))
 }
 const resolver:NativeFontResolver={providerId:'injoffice.pptx.operator-fonts',providerRevision:hash(Buffer.from(JSON.stringify([manifest,allowSubstitution,policy??null]))),resolve({run,manifest:requested}){
  if(JSON.stringify(requested)!==JSON.stringify(manifest))throw new Error('Operator font manifest binding changed')
  const selected=selectExplicitFontV1(manifest,run,allowSubstitution?policy:undefined)
  return selected?{status:'resolved',face:selected.face,attemptedFaceIds:[selected.face.faceId],decisions:[]}:{status:'refused',attemptedFaceIds:[],decisions:[{code:'font-not-found',message:`Configured font unavailable: ${run.font.families.join(', ')} / ${run.font.weight} / ${run.font.style}`,recoverable:false}]}
 },load(face){const resource=resources.get(face.faceId);if(!resource||resource.face.contentDigest!==face.contentDigest||resource.face.family!==face.family||resource.face.weight!==face.weight||resource.face.style!==face.style)throw new Error('Unresolved font resource');return {...resource,face:{...face},bytes:resource.bytes.slice()}}}
 return {manifest,resolver,outlines,resources,skipped,policyDigest:policy?hash(Buffer.from(JSON.stringify(policy))):undefined}
}

// uninspectable names the manifest faces that were dropped because HarfBuzz
// could not read them, so a missing family reads as "the manifest listed it and
// we could not load it" rather than "the manifest never listed it".
function uninspectable(fonts:{skipped:readonly string[]}):string{
 return fonts.skipped.length?`; ${fonts.skipped.length} operator face(s) were listed but could not be inspected: ${fonts.skipped.join(', ')}`:''
}

export async function compilePptxPreview(input:unknown):Promise<PptxPreview>{
 const request=object(input)
 if(request.workbook_chart_preview!==undefined&&typeof request.workbook_chart_preview!=='boolean')throw new Error('Workbook chart preview requires a boolean opt-in')
 if(request.workbook_chart_preview===true&&request.source_chart_preview===true)throw new Error('Chart preview modes are mutually exclusive')
 if((request.workbook_chart_preview===true)!==(request.workbook_chart_data!==undefined))throw new Error('Workbook chart data requires its exact opt-in')
 if(request.source_chart_preview!==undefined&&typeof request.source_chart_preview!=='boolean')throw new Error('Source chart preview requires a boolean opt-in')
 if(request.source_frame_autofit_preview!==undefined&&typeof request.source_frame_autofit_preview!=='boolean')throw new Error('Autofit preview requires a boolean opt-in')
 if(request.inherited_text_preview!==undefined&&typeof request.inherited_text_preview!=='boolean')throw new Error('Inherited text preview requires a boolean opt-in')
 if(request.font_substitution_preview!==undefined&&typeof request.font_substitution_preview!=='boolean')throw new Error('Font substitution preview requires a boolean opt-in')
 if(typeof request.package_sha256!=='string'||!/^[a-f0-9]{64}$/.test(request.package_sha256)||typeof request.font_manifest_path!=='string'||!Number.isSafeInteger(request.slide_index))throw new Error('Invalid source-bound preview input')
 assertNativePptx(request.deck)
 const deck=request.deck as NativePptxDeck
 if(deck.origin!=='parsed'||request.slide_index as number<0||request.slide_index as number>=deck.slides.length)throw new Error('Preview requires a parsed source slide')
 const workbookResolution=request.workbook_chart_preview===true?await workbookChartsForPreview(request.workbook_chart_data,deck,request.package_sha256):undefined
 const fonts=fontProviders(request.font_manifest_path,request.font_substitution_preview===true)
 const checkParagraphs=(paragraphs:readonly NativeParagraph[])=>{for(const paragraph of paragraphs)for(const run of paragraph.runs){if(![...fonts.resources.values()].some(r=>r.face.family===run.fontFamily&&r.face.weight===(run.bold?700:400)&&r.face.style===(run.italic?'italic':'normal')))throw new Error(`Exact operator font unavailable: ${run.fontFamily??'unresolved family'} / ${run.bold?'bold':'regular'} / ${run.italic?'italic':'normal'}${uninspectable(fonts)}`)}}
 const checkMarkerFonts=(paragraphs:readonly NativeParagraph[])=>{for(const p of paragraphs){if(!p.bulletFontFamily)continue;const run=p.runs[0];if(!run||![...fonts.resources.values()].some(r=>r.face.family===p.bulletFontFamily&&r.face.weight===(run.bold?700:400)&&r.face.style===(run.italic?'italic':'normal')))throw new Error(`Exact operator bullet font unavailable: ${p.bulletFontFamily}${uninspectable(fonts)}`)}}
 const checkElements=(elements:readonly NativeElement[])=>{for(const element of elements){if(element.kind==='text'||element.kind==='shape'){checkParagraphs(element.paragraphs);checkMarkerFonts(element.paragraphs)}if(element.kind==='group')checkElements(element.children);if(element.kind==='table')for(const row of element.table.rows)for(const cell of row)if(cell.paragraphs){checkParagraphs(cell.paragraphs);checkMarkerFonts(cell.paragraphs)}}}
 if(request.font_substitution_preview!==true)checkElements(deck.slides[request.slide_index as number]!.elements)
 // The v1 substitution evidence identifies text/shape runs, not table-cell
 // coordinates. Keep table font resolution exact until that source join exists.
 else {const tables=(elements:readonly NativeElement[])=>{for(const e of elements){if(e.kind==='group')tables(e.children);if(e.kind==='table')for(const row of e.table.rows)for(const cell of row)if(cell.paragraphs)checkParagraphs(cell.paragraphs)}};tables(deck.slides[request.slide_index as number]!.elements)}
 const countSourceFrames=(elements:readonly NativeElement[]):number=>elements.reduce((count,element)=>count+((element.kind==='text'||element.kind==='shape')&&element.textBody?.autoFit==='shape-source-frame'?1:0)+(element.kind==='group'?countSourceFrames(element.children):0),0)
 const sourceFrameAutoFitCount=countSourceFrames(deck.slides[request.slide_index as number]!.elements)
 // Authored paragraph spacing is resolved by the same Go-side inherited-text
 // cascade projection, so the worker re-checks that opt-in for it too.
 const countInherited=(elements:readonly NativeElement[]):number=>elements.reduce((n,e)=>n+(e.compatibility.diagnostics.some(d=>d.code==='pptx.source-inherited-text-approximate'||d.code==='pptx.paragraph-spacing-approximate')?1:0)+(e.kind==='group'?countInherited(e.children):0),0)
 const inheritedTextCount=countInherited(deck.slides[request.slide_index as number]!.elements)
 if(inheritedTextCount>0&&request.inherited_text_preview!==true)throw new Error('Inherited text requires explicit preview opt-in')
 if(sourceFrameAutoFitCount>0&&request.source_frame_autofit_preview!==true)throw new Error('Source-frame autofit requires explicit preview opt-in')
 // Authored normAutofit scale and single-column projections are Go-side read-only
 // approximations under the same autofit opt-in; the worker re-checks the gate.
 const countAuthoredAutoFit=(elements:readonly NativeElement[]):number=>elements.reduce((n,e)=>n+(e.compatibility.diagnostics.some(d=>d.code==='pptx.autofit-authored-scale-approximate'||d.code==='pptx.text-columns-approximate'||d.code==='pptx.text-warp-flattened-approximate'||d.code==='pptx.text-warp-approximate')?1:0)+(e.kind==='group'?countAuthoredAutoFit(e.children):0),0)
 if(countAuthoredAutoFit(deck.slides[request.slide_index as number]!.elements)>0&&request.source_frame_autofit_preview!==true)throw new Error('Authored autofit scale requires explicit preview opt-in')
 const tree=await compileNativePptxSlide(deck,request.slide_index as number,{workbookChartsPreview:workbookResolution?.charts,literalRadarPreview:request.source_chart_preview===true,literalStackedPreview:request.source_chart_preview===true,literalBubblePreview:request.source_chart_preview===true,literalAreaPreview:request.source_chart_preview===true,literalBarPreview:request.source_chart_preview===true,literalConnectedPreview:request.source_chart_preview===true,chartAxisLabelsPreview:request.source_chart_preview===true||request.workbook_chart_preview===true,inheritedTextPreview:request.inherited_text_preview===true,sourceFrameAutoFitPreview:request.source_frame_autofit_preview===true,lineLayoutPolicy:'max-run-natural-v1',maxGlyphs:20000,maxNodes:20000,textLayout:{glyphExtents:chartGlyphExtents(fonts.resources,fonts.outlines),manifest:fonts.manifest,resolver:fonts.resolver,shaper:createHarfBuzzTextShaperV1({sourceRevision:'pptx-preview-v1'}),defaults:{fontFamilies:[],fontSizeHundredthPt:1200,script:'Latn',language:'en-US',direction:'ltr'}}})
 const recording=createRecordingPaintSurface();paintSlideRenderTree(tree,recording)
 const substitutions:NonNullable<PptxPreview['font_substitutions']>=[]
 const root:Extract<PreviewNode,{kind:'group'}>={kind:'group',transform:[1,0,0,1,0,0],children:[]}
 const stack=[root],diagnostics=tree.diagnostics.map(d=>`${d.code}: ${d.message}`)
 const affineBudget=new SourceAffineBudget()
 let glyphs=0
 const resources=new Map<string,NativeDocxPagePaintMediaAssetV1>()
 for(const command of recording.finish()){
  const current=stack[stack.length-1]!
  switch(command.kind){
   case 'save':{const group:typeof root={kind:'group',transform:[1,0,0,1,0,0],children:[]};current.children.push(group);stack.push(group);break}
   case 'restore':if(stack.length<2)throw new Error('Paint stack underflow');stack.pop();break
   case 'transform':{current.transform=[...renderTransformMatrix(command.transform,affineBudget)];break}
   case 'clipRect':current.clip=command.rect;break
   case 'clipRoundRect':current.clip={...command.rect,radius:command.radiusEmu};break
   case 'clipPath':current.clip={...command.rect,d:command.path.map(pathPart).join(' ')};break
   case 'path':{
    if(command.headArrow&&!command.headEnd||command.tailArrow&&!command.tailEnd){diagnostics.push('Boolean-only arrowheads lack source type/dimensions and remain unqualified');current.children.push({kind:'placeholder',rect:{x:0,y:0,cx:300000,cy:100000},label:'Arrowhead unavailable'});break}
    const stroke=command.stroke
    const paint={fill:previewColor(command.fill??'none'),...(stroke?previewStroke(stroke):{})}
    const first=command.path[0]
    if(first?.kind==='rect'||first?.kind==='roundRect')current.children.push({kind:'rect',rect:first.rect,radius:first.kind==='roundRect'?first.radiusEmu:0,...paint})
    else if(first?.kind==='ellipse')current.children.push({kind:'ellipse',rect:first.rect,...paint})
    else {
     let d=command.path.map(pathPart).join(' ')
     const typed=command.headEnd?.type!=='none'&&command.headEnd||command.tailEnd?.type!=='none'&&command.tailEnd
     let shaft:ReturnType<typeof previewConnectorShaft>|undefined
     const arrows:PreviewNode[]=[]
     // Arrow-v1 endpoint geometry is element-scoped: a shaft or arrowhead that
     // cannot be qualified paints this element as a placeholder and never
     // rejects the rest of the slide.
     try{
      const arrowWidth=stroke?previewArrowStrokeBase(stroke.widthEmu):0
      shaft=(command.headEnd||command.tailEnd)&&stroke?previewConnectorShaft(command.path,command.headEnd,command.tailEnd,arrowWidth):undefined
      if(typed){
       if(!shaft||!stroke)throw new Error('Typed arrows require a source-bound stroked connector')
       for(const [end,terminal] of [[command.headEnd,shaft.head],[command.tailEnd,shaft.tail]] as const){if(!end)continue;const arrow=previewArrow(end,terminal.tip,terminal.direction,arrowWidth,stroke.color);if(arrow)arrows.push(arrow)}
      }
     }catch(error){
      diagnostics.push(`connector.arrowUnavailable: ${error instanceof Error?error.message:'arrow-v1 endpoint geometry unavailable'}; element painted as a placeholder`)
      current.children.push({kind:'placeholder',rect:{x:0,y:0,cx:300000,cy:100000},label:'Connector arrow unavailable'})
      break
     }
     if(shaft)d=shaft.d
     current.children.push({kind:'path',d,...paint},...arrows)
     if(typed&&shaft){
      diagnostics.push('arrow.deterministicGeometry: InjOffice arrow-v1 uses source type and named widths/lengths (2/3/5 × max(stroke, 2pt hairline); omitted = medium), not Office-equivalent geometry')
      if(!shaft.straight)diagnostics.push('connector.presetShaftPreview: arrow-v1 endpoints follow the terminal tangents of the evaluated connector-preset path (pptx.connector-preset-preview); not Office-equivalent geometry')
     }
     if(shaft?.skippedInsets.length)diagnostics.push(`connector.shaftInsetSkipped: ${shaft.skippedInsets.join('/')} terminal segment shorter than the arrow-v1 inset; shaft left untrimmed under the arrowhead`)
    }
    break
   }
   case 'glyphRun':{
    const run=command.run,provider=fonts.outlines.get(run.faceId??'')
    if(run.fontSelection&&run.fontSelection.resolution!=='exact'){
     if(!fonts.policyDigest||request.font_substitution_preview!==true||run.sourceRole==='paragraphBullet'||!run.faceId||!run.contentDigest)throw new Error('Unattested font substitution')
     const evidence={source_id:run.sourceElementId,paragraph_index:run.paragraphIndex,run_index:run.runIndex,source_family:run.fontSelection.sourceFamily,selected_family:run.fontSelection.selectedFamily,face_id:run.faceId,font_digest:run.contentDigest}
     if(!substitutions.some(s=>JSON.stringify(s)===JSON.stringify(evidence)))substitutions.push(evidence)
    }
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
 if(request.source_chart_preview===true){result.source_chart_preview=true;result.chart_axis_layout_policy='supplied-outline-margins-v1'}
 if(request.workbook_chart_preview===true){result.workbook_chart_preview=true;result.chart_axis_layout_policy='supplied-outline-margins-v1';for(const refusal of workbookResolution!.refusals)result.diagnostics.push(`Workbook chart ${refusal.objectId}: ${refusal.reason}`.slice(0,2048))}
 if(sourceFrameAutoFitCount>0)result.source_frame_autofit_count=sourceFrameAutoFitCount
 if(inheritedTextCount>0){result.inherited_text_preview_count=inheritedTextCount;result.inherited_text_policy='source-latin-inheritance-approximate-v1'}
 if(substitutions.length){result.font_substitutions=substitutions;result.font_substitution_policy=EXPLICIT_FONT_POLICY_V1;result.font_substitution_policy_sha256=fonts.policyDigest}
 return decodePptxPreview(result)
}
function pathPart(p:RenderPathCommand):string {switch(p.kind){case 'moveTo':return `M${p.x} ${p.y}`;case 'lineTo':return `L${p.x} ${p.y}`;case 'quadBezierTo':return `Q${p.x1} ${p.y1} ${p.x} ${p.y}`;case 'cubicBezierTo':return `C${p.x1} ${p.y1} ${p.x2} ${p.y2} ${p.x} ${p.y}`;case 'arcTo':return `A${p.rx} ${p.ry} 0 ${p.largeArc?1:0} ${p.clockwise?1:0} ${p.x} ${p.y}`;case 'close':return 'Z';default:throw new Error('Unmodeled path command')}}
