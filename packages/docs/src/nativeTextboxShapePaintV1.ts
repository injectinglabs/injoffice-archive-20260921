import {planTextboxWrap,type NativeTextboxWrapClusterV1,type NativeTextboxWrapPaintV1} from './nativeTextboxWrappingV1.js'
import type {NativeTextboxLinePaintV1} from './nativeTextboxHardBreaksV1.js'
/** Node-only deterministic textbox-local compiler. No platform font lookup. */
import {createHarfBuzzTextShaperV1,createHarfBuzzOutlineProviderV1,inspectHarfBuzzFontMetricsV1,HARFBUZZ_SHAPER_CONFIG_REVISION,type HarfBuzzOutlineCommandV1} from '@injoffice/font-metrics/harfbuzz'
import {NATIVE_TEXT_LAYOUT_VERSION,type FontResource,type TextRunInput,type ShapedSegment} from '@injoffice/font-metrics/layout'
import {sha256} from '@noble/hashes/sha2.js'
import {bytesToHex} from '@noble/hashes/utils.js'
import {decodeNativeDocxTextboxGeometryV1,decodeNativeDocxTextboxShapePaintV1,nativeTextboxGeometryDigestV1,type NativeDocxTextboxShapePaintV1} from './nativeTextboxGeometryPreviewV1.js'

// This narrow name/regular-face check runs after canonical SFNT validation.
// Only unambiguous Windows Unicode English family/subfamily records qualify.
function fontMatches(bytes:Uint8Array,family:string):boolean {
 const v=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength),tables=new Map<string,number>()
 for(let i=0;i<v.getUint16(4);i++){const p=12+i*16;tables.set(String.fromCharCode(...bytes.subarray(p,p+4)),v.getUint32(p+8))}
 const name=tables.get('name'),os2=tables.get('OS/2'),head=tables.get('head')
 if(name===undefined||os2===undefined||head===undefined||v.getUint16(os2+4)!==400||v.getUint16(os2+6)!==5||(v.getUint16(os2+62)&0x21)!==0||(v.getUint16(head+44)&3)!==0)return false
 const families=new Set<string>(),styles=new Set<string>(),count=v.getUint16(name+2),strings=name+v.getUint16(name+4)
 if(count>4096)return false
 for(let i=0;i<count;i++){
  const p=name+6+i*12,platform=v.getUint16(p),encoding=v.getUint16(p+2),lang=v.getUint16(p+4),id=v.getUint16(p+6),len=v.getUint16(p+8),off=strings+v.getUint16(p+10)
  if(platform!==3||(encoding!==1&&encoding!==10)||lang!==0x409||![1,2,16,17].includes(id))continue
  if(len===0||len>512||len%2||off+len>bytes.length)return false
  let text='';for(let j=0;j<len;j+=2)text+=String.fromCharCode(v.getUint16(off+j))
  if(id===1||id===16)families.add(text);else styles.add(text)
 }
 return families.size===1&&families.has(family)&&styles.size>0&&[...styles].every(s=>s==='Regular'||s==='Book')
}

/** Compile one source-qualified rectangle using the actual supplied regular
 * font bytes. Unsupported source, font identity, or any overflow stays omitted. */
export function compileNativeDocxTextboxShapeV1(source:unknown,evidence:unknown,index:number,fontBytes:Uint8Array):NativeDocxTextboxShapePaintV1 {
 const joined=decodeNativeDocxTextboxGeometryV1(source,evidence),item=joined.items[index]
 if(!Number.isSafeInteger(index)||!item)throw new TypeError('Unknown textbox index')
 if(!(fontBytes instanceof Uint8Array)||fontBytes.byteLength<1||fontBytes.byteLength>16*1024*1024)throw new TypeError('Font bytes must be 1..16 MiB')
 const bytes=Uint8Array.from(fontBytes),digest=('sha256:'+bytesToHex(sha256(bytes))) as `sha256:${string}`
 const result:NativeDocxTextboxShapePaintV1={protocol:'injoffice.docx.textbox-shape-paint',version:1,package_sha256:item.owner.package_sha256,diagnostic_id:item.owner.diagnostic_id,input_sha256:nativeTextboxGeometryDigestV1(item),font_sha256:digest,status:'omitted',reason:'unsupported-shape-source',paths:[],width_millipoints:0,height_millipoints:0,line_width_millipoints:0,fill_rgb:'none',line_rgb:'none',text_rgb:'000000',shaper_revision:HARFBUZZ_SHAPER_CONFIG_REVISION}
 const g=item.geometry;if(!g)return result
 try{
  const request={bytes,contentDigest:digest},metrics=inspectHarfBuzzFontMetricsV1(request)
  if(!fontMatches(bytes,g.font_family)){result.reason='font-family-or-regular-face-mismatch';return result}
  const font:FontResource={bytes,metrics,face:{faceId:digest,family:g.font_family,weight:400,style:'normal',stretch:100,sourceKind:'host',resourceId:digest,contentDigest:digest,resolution:'exact',matchedFamily:g.font_family}}
  const hard=item.hard_break_layout,wrap=item.wrap_layout
  const emu=(n:number)=>n*10/127,w=emu(g.width_emu),h=emu(g.height_emu),[left,top,right,bottom]=g.insets_emu.map(emu) as [number,number,number,number]
  const makeRun=(text:string):TextRunInput=>({version:NATIVE_TEXT_LAYOUT_VERSION,text,fontSizeMilliPoints:g.font_size_half_points*500,font:{families:[g.font_family],weight:400,style:'normal',stretch:100},direction:'ltr',script:'Latn',language:'en-US',features:[{tag:'kern',value:0}]})
  let whole:ShapedSegment|undefined,wrapPaint:NativeTextboxWrapPaintV1|undefined,wrapPlan:ReturnType<typeof planTextboxWrap>|undefined
  if(wrap){
   const run=makeRun(item.owner.paragraphs[0]!),shaped=createHarfBuzzTextShaperV1({sourceRevision:'injoffice.textbox-shape-v1'}).shape({run,startUtf16:0,endUtf16:run.text.length,font})
   if('status'in shaped){result.reason='font-shaping-refused';return result}
   whole=shaped
   const clusters:NativeTextboxWrapClusterV1[]=shaped.clusters.map(c=>({start_utf16:c.startUtf16,end_utf16:c.endUtf16,advance_millipoints:c.advanceInlineMilliPoints,unsafe_to_break:c.unsafeToBreak===true,glyph_start:c.glyphStart,glyph_end:c.glyphEnd}))
   wrapPlan=planTextboxWrap(run.text,clusters,w-left-right,wrap.policy)
   if(clusters.at(-1)!.glyph_end!==shaped.glyphs.length)throw new RangeError('wrap-glyph-coverage')
   wrapPaint={clusters,natural_height_millipoints:shaped.metrics.lineHeightMilliPoints,ascent_millipoints:shaped.metrics.ascentMilliPoints,line_gap_millipoints:0,lines:[]}
  }
  const sourceLines=wrapPlan?wrapPlan.map(l=>item.owner.paragraphs[0]!.slice(l.start_utf16,l.end_utf16)):hard?hard.lines.map(l=>item.owner.paragraphs[0]!.slice(l.start_utf16,l.end_utf16)):[item.owner.paragraphs[0]!]
  const paths:string[]=[],provider=createHarfBuzzOutlineProviderV1(request),linePaint:NativeTextboxLinePaintV1={natural_height_millipoints:0,ascent_millipoints:0,line_gap_millipoints:0,lines:[]}
  let budget=0,glyphCount=0
  for(let ordinal=0;ordinal<sourceLines.length;ordinal++){
  const run=makeRun(sourceLines[ordinal]!),planned=wrapPlan?.[ordinal]
  const shaped=whole&&planned?{...whole,advanceInlineMilliPoints:planned.advance_millipoints,glyphs:whole.glyphs.slice(whole.clusters[planned.cluster_start]!.glyphStart,whole.clusters[planned.cluster_end-1]!.glyphEnd)}:createHarfBuzzTextShaperV1({sourceRevision:'injoffice.textbox-shape-v1'}).shape({run,startUtf16:0,endUtf16:run.text.length,font})
  if('status'in shaped){result.reason='font-shaping-refused';return result}
  if(shaped.advanceInlineMilliPoints>w-left-right||shaped.metrics.lineHeightMilliPoints>h-top-bottom){result.reason='text-overflow';return result}
  const explicit=hard??wrap,step=explicit?explicit.line_step_twips*50:shaped.metrics.lineHeightMilliPoints,center=explicit?(step-shaped.metrics.lineHeightMilliPoints)/2:0
  if(explicit&&(shaped.metrics.lineGapMilliPoints!==0||center<0||!Number.isInteger(center)||step*sourceLines.length>h-top-bottom)){result.reason='exact-line-layout-refused';return result}
  if(hard&&ordinal>0&&(linePaint.natural_height_millipoints!==shaped.metrics.lineHeightMilliPoints||linePaint.ascent_millipoints!==shaped.metrics.ascentMilliPoints)){result.reason='inconsistent-line-metrics';return result}
  const lineTop=top+ordinal*step,lineBottom=explicit?lineTop+step:h-bottom,baseline=lineTop+center+shaped.metrics.ascentMilliPoints,pathStart=paths.length
  let x=left
  for(const glyph of shaped.glyphs){
   if(++glyphCount>16384)throw new RangeError('glyph-budget')
   const outline=provider.outline(glyph.glyphId),scale=run.fontSizeMilliPoints/outline.units_per_em
   const point=(a:number,b:number):string=>{
    const px=Math.round(x+glyph.offsetXMilliPoints+a*scale),py=Math.round(baseline-glyph.offsetYMilliPoints-b*scale)
    if(!Number.isSafeInteger(px)||!Number.isSafeInteger(py)||px<left||px>w-right||py<lineTop||py>lineBottom)throw new RangeError('glyph-overflow')
    return `${px} ${py}`
   }
   const command=(c:HarfBuzzOutlineCommandV1):string=>{
    switch(c.kind){case 'close_path':return 'Z';case 'move_to':return 'M'+point(c.x,c.y);case 'line_to':return 'L'+point(c.x,c.y);case 'quadratic_to':return 'Q'+point(c.control_x,c.control_y)+' '+point(c.x,c.y);case 'cubic_to':return 'C'+point(c.control_1_x,c.control_1_y)+' '+point(c.control_2_x,c.control_2_y)+' '+point(c.x,c.y)}
   }
   const path=outline.path.map(command).join(' ');budget+=path.length;if(path.length>1000000||budget>8000000)throw new RangeError('glyph-budget')
   if(path)paths.push(path);x+=glyph.advanceXMilliPoints
  }
  if(wrapPaint&&planned)wrapPaint.lines.push({...planned,baseline_millipoints:baseline,path_start:pathStart,path_end:paths.length})
  if(hard){const sourceLine=hard.lines[ordinal]!;linePaint.natural_height_millipoints=shaped.metrics.lineHeightMilliPoints;linePaint.ascent_millipoints=shaped.metrics.ascentMilliPoints;linePaint.lines.push({ordinal,start_utf16:sourceLine.start_utf16,end_utf16:sourceLine.end_utf16,baseline_millipoints:baseline,path_start:pathStart,path_end:paths.length})}
  }
  if(hard)result.line_layout=linePaint
  if(wrapPaint)result.wrap_paint=wrapPaint
  Object.assign(result,{status:'supported',reason:'',paths,width_millipoints:w,height_millipoints:h,line_width_millipoints:emu(g.line_width_emu),fill_rgb:g.fill_rgb,line_rgb:g.line_rgb,text_rgb:g.text_rgb})
  return decodeNativeDocxTextboxShapePaintV1(source,joined,index,result,digest)
 }catch(error){delete result.line_layout;delete result.wrap_paint;Object.assign(result,{status:'omitted',paths:[],width_millipoints:0,height_millipoints:0,line_width_millipoints:0,fill_rgb:'none',line_rgb:'none',text_rgb:'000000'});result.reason=error instanceof RangeError&&(error.message==='glyph-overflow'||error.message.startsWith('wrap-'))?error.message:'font-or-outline-refused';return result}
}
