import {classifyNativeOfficeLineBreakRanges} from '@injoffice/font-metrics/layout'
import type {NativeDocxSourceAnchorV1} from './nativeContract.js'
import type {NativeDocxTextboxV1} from './nativeTextboxInventoryV1.js'

export interface NativeTextboxWrapLayoutV1 {policy:'ascii-space-greedy-v1';body_properties_anchor:NativeDocxSourceAnchorV1;paragraph_anchor:NativeDocxSourceAnchorV1;run_anchor:NativeDocxSourceAnchorV1;text_anchor:NativeDocxSourceAnchorV1;spacing_anchor:NativeDocxSourceAnchorV1;line_step_twips:number}
export interface NativeTextboxWrapClusterV1 {start_utf16:number;end_utf16:number;advance_millipoints:number;unsafe_to_break:boolean;glyph_start:number;glyph_end:number}
export interface NativeTextboxWrapLineV1 {ordinal:number;start_utf16:number;end_utf16:number;cluster_start:number;cluster_end:number;advance_millipoints:number}
export interface NativeTextboxWrapPaintV1 {clusters:NativeTextboxWrapClusterV1[];natural_height_millipoints:number;ascent_millipoints:number;line_gap_millipoints:0;lines:(NativeTextboxWrapLineV1&{baseline_millipoints:number;path_start:number;path_end:number})[]}
function keys(v:unknown,wanted:string[]):asserts v is Record<string,unknown>{if(!v||typeof v!=='object'||Array.isArray(v)||Object.keys(v).sort().join(',')!==wanted.sort().join(','))throw new TypeError('Invalid textbox wrap fields')}
function anchor(v:unknown,parent:NativeDocxSourceAnchorV1,path?:string):NativeDocxSourceAnchorV1{
 keys(v,['part_name','path','start_byte','end_byte','xml_sha256'])
 if(v.part_name!==parent.part_name||typeof v.path!=='string'||(path?v.path!==path:!v.path.startsWith(parent.path+'/'))||!Number.isSafeInteger(v.start_byte)||!Number.isSafeInteger(v.end_byte)||Number(v.start_byte)<=parent.start_byte||Number(v.end_byte)>=parent.end_byte||Number(v.start_byte)>=Number(v.end_byte)||typeof v.xml_sha256!=='string'||!/^sha256:[0-9a-f]{64}$/.test(v.xml_sha256))throw new TypeError('Invalid textbox wrap anchor')
 return v as unknown as NativeDocxSourceAnchorV1
}
/** Input already passed bounded own-data copying and owner joins. */
export function decodeTextboxWrapSource(v:unknown,owner:NativeDocxTextboxV1):NativeTextboxWrapLayoutV1{
 keys(v,['policy','body_properties_anchor','paragraph_anchor','run_anchor','text_anchor','spacing_anchor','line_step_twips'])
 if(v.policy!=='ascii-space-greedy-v1'||!Number.isSafeInteger(v.line_step_twips)||Number(v.line_step_twips)<1||Number(v.line_step_twips)>25600||owner.paragraphs.length!==1||owner.paragraphs[0]!.length>4096||!/^[A-Za-z0-9]+(?: [A-Za-z0-9]+)*$/.test(owner.paragraphs[0]!))throw new TypeError('Invalid textbox wrap profile')
 const p=anchor(v.paragraph_anchor,owner.anchor),r=anchor(v.run_anchor,p,p.path+'/w:r[1]'),t=anchor(v.text_anchor,r,r.path+'/w:t[1]'),spacing=anchor(v.spacing_anchor,p,p.path+'/w:pPr[1]/w:spacing[1]'),body=anchor(v.body_properties_anchor,owner.anchor)
 const root=owner.anchor.path.match(/^(.*\/w:drawing\[[1-9][0-9]*\])/u)?.[1],suffix=root&&p.path.slice(root.length)
 if(!root||!/^\/(?:wp|ns[0-9a-f]{8}):inline\[1\]\/(?:a|ns[0-9a-f]{8}):graphic\[1\]\/(?:a|ns[0-9a-f]{8}):graphicData\[1\]\/(?:wps|ns[0-9a-f]{8}):wsp\[1\]\/(?:wps|ns[0-9a-f]{8}):txbx\[1\]\/w:txbxContent\[1\]\/w:p\[1\]$/.test(suffix!)||spacing.end_byte>=r.start_byte||body.start_byte<=p.end_byte)throw new TypeError('Invalid wrap source order')
 const shapePath=p.path.replace(/\/(?:wps|ns[0-9a-f]{8}):txbx\[1\]\/w:txbxContent\[1\]\/w:p\[1\]$/,'')
 if(!new RegExp('^'+shapePath.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')+'/(?:wps|ns[0-9a-f]{8}):bodyPr\\[1\\]$').test(body.path))throw new TypeError('Invalid wrap body properties')
 return {policy:'ascii-space-greedy-v1',body_properties_anchor:body,paragraph_anchor:p,run_anchor:r,text_anchor:t,spacing_anchor:spacing,line_step_twips:Number(v.line_step_twips)}
}
/** Greedy complete-cluster selection, with spaces retained on their source line. */
export function planTextboxWrap(text:string,clusters:readonly NativeTextboxWrapClusterV1[],width:number):NativeTextboxWrapLineV1[]{
 if(!Number.isSafeInteger(width)||width<1||text.length>4096||!/^[A-Za-z0-9]+(?: [A-Za-z0-9]+)*$/.test(text)||!clusters.length||clusters.length>4096)throw new TypeError('Textbox wrap budget')
 let end=0,glyphEnd=0
 for(const c of clusters){keys(c,['start_utf16','end_utf16','advance_millipoints','unsafe_to_break','glyph_start','glyph_end']);if(c.glyph_start!==glyphEnd||!Number.isSafeInteger(c.glyph_end)||c.glyph_end<=glyphEnd||c.glyph_end>16384||c.start_utf16!==end||!Number.isSafeInteger(c.end_utf16)||c.end_utf16<=end||c.end_utf16>text.length||!/^(?:[A-Za-z0-9]+| )$/.test(text.slice(end,c.end_utf16))||!Number.isSafeInteger(c.advance_millipoints)||c.advance_millipoints<0||c.advance_millipoints>10000000||typeof c.unsafe_to_break!=='boolean')throw new TypeError('Invalid wrap cluster coverage');end=c.end_utf16;glyphEnd=c.glyph_end}
 if(end!==text.length)throw new TypeError('Incomplete wrap cluster coverage')
 const lines:NativeTextboxWrapLineV1[]=[];let start=0,work=0
 while(start<clusters.length){
  if(lines.length>=16)throw new RangeError('wrap-line-budget')
  let advance=0,last=-1,lastAdvance=0,i=start
  for(;i<clusters.length;i++){
   if(++work>65536)throw new RangeError('wrap-work-budget')
   const c=clusters[i]!,next=clusters[i+1],sum=advance+c.advance_millipoints
   if(sum>width)break;advance=sum
   if(i===clusters.length-1){last=i+1;lastAdvance=advance;break}
   if(text.slice(c.start_utf16,c.end_utf16)===' '&&!c.unsafe_to_break&&!next!.unsafe_to_break&&classifyNativeOfficeLineBreakRanges(text,c.start_utf16,c.end_utf16,text,next!.start_utf16,next!.end_utf16)==='allowed'){last=i+1;lastAdvance=advance}
  }
  if(last<=start)throw new RangeError('wrap-word-or-space-overflow')
  lines.push({ordinal:lines.length,start_utf16:clusters[start]!.start_utf16,end_utf16:clusters[last-1]!.end_utf16,cluster_start:start,cluster_end:last,advance_millipoints:lastAdvance});start=last
 }
 return lines
}
export function decodeTextboxWrapPaint(v:unknown,source:NativeTextboxWrapLayoutV1,text:string,width:number,top:number,bottom:number,pathCount:number):NativeTextboxWrapPaintV1{
 keys(v,['clusters','natural_height_millipoints','ascent_millipoints','line_gap_millipoints','lines'])
 if(!Array.isArray(v.clusters)||!Array.isArray(v.lines))throw new TypeError('Invalid wrap arrays')
 const planned=planTextboxWrap(text,v.clusters,width),height=Number(v.natural_height_millipoints),ascent=Number(v.ascent_millipoints),step=source.line_step_twips*50,center=(step-height)/2
 if(!Number.isSafeInteger(v.natural_height_millipoints)||height<=0||!Number.isSafeInteger(v.ascent_millipoints)||ascent<=0||ascent>height||v.line_gap_millipoints!==0||center<0||!Number.isInteger(center)||planned.length*step>bottom-top||v.lines.length!==planned.length)throw new TypeError('Invalid wrapped line metrics')
 let pathEnd=0
 for(let i=0;i<planned.length;i++){
  const l=v.lines[i],p=planned[i]!;keys(l,['ordinal','start_utf16','end_utf16','cluster_start','cluster_end','advance_millipoints','baseline_millipoints','path_start','path_end'])
  if(Object.entries(p).some(([k,x])=>l[k]!==x)||l.baseline_millipoints!==top+center+ascent+i*step||l.path_start!==pathEnd||!Number.isSafeInteger(l.path_end)||Number(l.path_end)<=pathEnd||Number(l.path_end)-pathEnd>v.clusters[p.cluster_end-1]!.glyph_end-v.clusters[p.cluster_start]!.glyph_start||Number(l.path_end)>pathCount)throw new TypeError('Invalid wrapped line paint coverage');pathEnd=Number(l.path_end)
 }
 if(pathEnd!==pathCount)throw new TypeError('Incomplete wrapped paths')
 return v as unknown as NativeTextboxWrapPaintV1
}
