/** Node-only deterministic textbox-local compiler. No platform font lookup. */
import {createHarfBuzzTextShaperV1,createHarfBuzzOutlineProviderV1,inspectHarfBuzzFontMetricsV1,HARFBUZZ_SHAPER_CONFIG_REVISION,type HarfBuzzOutlineCommandV1} from '@injoffice/font-metrics/harfbuzz'
import {NATIVE_TEXT_LAYOUT_VERSION,type FontResource,type TextRunInput} from '@injoffice/font-metrics/layout'
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
  const run:TextRunInput={version:NATIVE_TEXT_LAYOUT_VERSION,text:item.owner.paragraphs[0]!,fontSizeMilliPoints:g.font_size_half_points*500,font:{families:[g.font_family],weight:400,style:'normal',stretch:100},direction:'ltr',script:'Latn',language:'en-US',features:[{tag:'kern',value:0}]}
  const shaped=createHarfBuzzTextShaperV1({sourceRevision:'injoffice.textbox-shape-v1'}).shape({run,startUtf16:0,endUtf16:run.text.length,font})
  if('status'in shaped){result.reason='font-shaping-refused';return result}
  const emu=(n:number)=>Math.round(n*10/127),w=emu(g.width_emu),h=emu(g.height_emu),[left,top,right,bottom]=g.insets_emu.map(emu) as [number,number,number,number]
  if(shaped.advanceInlineMilliPoints>w-left-right||shaped.metrics.lineHeightMilliPoints>h-top-bottom){result.reason='text-overflow';return result}
  const baseline=top+shaped.metrics.ascentMilliPoints,provider=createHarfBuzzOutlineProviderV1(request),paths:string[]=[]
  let x=left,budget=0
  for(const glyph of shaped.glyphs){
   const outline=provider.outline(glyph.glyphId),scale=run.fontSizeMilliPoints/outline.units_per_em
   const point=(a:number,b:number):string=>{
    const px=Math.round(x+glyph.offsetXMilliPoints+a*scale),py=Math.round(baseline-glyph.offsetYMilliPoints-b*scale)
    if(!Number.isSafeInteger(px)||!Number.isSafeInteger(py)||px<left||px>w-right||py<top||py>h-bottom)throw new RangeError('glyph-overflow')
    return `${px} ${py}`
   }
   const command=(c:HarfBuzzOutlineCommandV1):string=>{
    switch(c.kind){case 'close_path':return 'Z';case 'move_to':return 'M'+point(c.x,c.y);case 'line_to':return 'L'+point(c.x,c.y);case 'quadratic_to':return 'Q'+point(c.control_x,c.control_y)+' '+point(c.x,c.y);case 'cubic_to':return 'C'+point(c.control_1_x,c.control_1_y)+' '+point(c.control_2_x,c.control_2_y)+' '+point(c.x,c.y)}
   }
   const path=outline.path.map(command).join(' ');budget+=path.length;if(path.length>1000000||budget>8000000)throw new RangeError('glyph-budget')
   if(path)paths.push(path);x+=glyph.advanceXMilliPoints
  }
  Object.assign(result,{status:'supported',reason:'',paths,width_millipoints:w,height_millipoints:h,line_width_millipoints:emu(g.line_width_emu),fill_rgb:g.fill_rgb,line_rgb:g.line_rgb,text_rgb:g.text_rgb})
  return decodeNativeDocxTextboxShapePaintV1(source,joined,index,result,digest)
 }catch(error){Object.assign(result,{status:'omitted',paths:[],width_millipoints:0,height_millipoints:0,line_width_millipoints:0,fill_rgb:'none',line_rgb:'none',text_rgb:'000000'});result.reason=error instanceof RangeError&&error.message==='glyph-overflow'?'glyph-overflow':'font-or-outline-refused';return result}
}
