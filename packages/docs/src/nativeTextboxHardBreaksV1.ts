import type {NativeDocxSourceAnchorV1} from './nativeContract.js'
import type {NativeDocxTextboxV1} from './nativeTextboxInventoryV1.js'

export interface NativeTextboxHardBreakLineV1 {ordinal:number;text_anchor:NativeDocxSourceAnchorV1;break_before_anchor:NativeDocxSourceAnchorV1|null;start_utf16:number;end_utf16:number}
export interface NativeTextboxHardBreakLayoutV1 {paragraph_anchor:NativeDocxSourceAnchorV1;run_anchor:NativeDocxSourceAnchorV1;spacing_anchor:NativeDocxSourceAnchorV1;line_step_twips:number;lines:NativeTextboxHardBreakLineV1[]}
export interface NativeTextboxLinePaintV1 {natural_height_millipoints:number;ascent_millipoints:number;line_gap_millipoints:0;lines:{ordinal:number;start_utf16:number;end_utf16:number;baseline_millipoints:number;path_start:number;path_end:number}[]}
function keys(v:unknown,wanted:string[]):asserts v is Record<string,unknown>{if(!v||typeof v!=='object'||Array.isArray(v)||Object.keys(v).sort().join(',')!==wanted.sort().join(','))throw new TypeError('Invalid hard-break fields')}
function anchor(v:unknown,parent:NativeDocxSourceAnchorV1,path?:string):NativeDocxSourceAnchorV1{
 keys(v,['part_name','path','start_byte','end_byte','xml_sha256'])
 if(v.part_name!==parent.part_name||typeof v.path!=='string'||(path?v.path!==path:!v.path.startsWith(parent.path+'/'))||!Number.isSafeInteger(v.start_byte)||!Number.isSafeInteger(v.end_byte)||Number(v.start_byte)<=parent.start_byte||Number(v.end_byte)>=parent.end_byte||Number(v.start_byte)>=Number(v.end_byte)||typeof v.xml_sha256!=='string'||!/^sha256:[0-9a-f]{64}$/.test(v.xml_sha256))throw new TypeError('Invalid hard-break anchor')
 return v as unknown as NativeDocxSourceAnchorV1
}
/** Called only after bounded own-data copying and diagnostic identity joins. */
export function decodeTextboxHardBreaks(value:unknown,owner:NativeDocxTextboxV1):NativeTextboxHardBreakLayoutV1{
 keys(value,['paragraph_anchor','run_anchor','spacing_anchor','line_step_twips','lines'])
 if(!Number.isSafeInteger(value.line_step_twips)||Number(value.line_step_twips)<1||Number(value.line_step_twips)>25600||!Array.isArray(value.lines)||value.lines.length<2||value.lines.length>16||owner.paragraphs.length!==1)throw new TypeError('Invalid hard-break limits')
 const p=anchor(value.paragraph_anchor,owner.anchor),r=anchor(value.run_anchor,p,p.path+'/w:r[1]'),spacing=anchor(value.spacing_anchor,p,p.path+'/w:pPr[1]/w:spacing[1]')
 const root=owner.anchor.path.match(/^(.*\/w:drawing\[[1-9][0-9]*\])/u)?.[1]
 const suffix=root&&p.path.slice(root.length)
 if(!root||!/^\/(?:wp|ns[0-9a-f]{8}):(?:inline|anchor)\[1\]\/(?:a|ns[0-9a-f]{8}):graphic\[1\]\/(?:a|ns[0-9a-f]{8}):graphicData\[1\]\/(?:wps|ns[0-9a-f]{8}):wsp\[1\]\/(?:wps|ns[0-9a-f]{8}):txbx\[1\]\/w:txbxContent\[1\]\/w:p\[1\]$/.test(suffix!)||spacing.end_byte>=r.start_byte)throw new TypeError('Invalid hard-break paragraph sequence')
 const text=owner.paragraphs[0]!,lines:NativeTextboxHardBreakLineV1[]=[]
 let offset=0,end=r.start_byte
 for(let i=0;i<value.lines.length;i++){
  const line=value.lines[i];keys(line,['ordinal','text_anchor','break_before_anchor','start_utf16','end_utf16'])
  const t=anchor(line.text_anchor,r,r.path+`/w:t[${i+1}]`)
  let br:NativeDocxSourceAnchorV1|null=null
  if(i===0){if(line.break_before_anchor!==null)throw new TypeError('Leading hard break')}
  else {br=anchor(line.break_before_anchor,r,r.path+`/w:br[${i}]`);if(br.start_byte<end||br.end_byte>t.start_byte||text[offset]!=='\n')throw new TypeError('Hard-break source order');offset++}
  if(line.ordinal!==i||line.start_utf16!==offset||!Number.isSafeInteger(line.end_utf16)||Number(line.end_utf16)<=offset||Number(line.end_utf16)>text.length||t.start_byte<end||!/^[\x20-\x7e]+$/.test(text.slice(offset,Number(line.end_utf16))))throw new TypeError('Hard-break text coverage')
  offset=Number(line.end_utf16);end=t.end_byte;lines.push({ordinal:i,text_anchor:t,break_before_anchor:br,start_utf16:line.start_utf16 as number,end_utf16:offset})
 }
 if(offset!==text.length||text.length>4096)throw new TypeError('Incomplete hard-break coverage')
 return {paragraph_anchor:p,run_anchor:r,spacing_anchor:spacing,line_step_twips:Number(value.line_step_twips),lines}
}
export function decodeTextboxLinePaint(value:unknown,source:NativeTextboxHardBreakLayoutV1,top:number,bottom:number,pathCount:number,text:string):NativeTextboxLinePaintV1{
 keys(value,['natural_height_millipoints','ascent_millipoints','line_gap_millipoints','lines'])
 const height=Number(value.natural_height_millipoints),ascent=Number(value.ascent_millipoints),step=source.line_step_twips*50,center=(step-height)/2
 if(!Number.isSafeInteger(value.natural_height_millipoints)||height<=0||!Number.isSafeInteger(value.ascent_millipoints)||ascent<=0||ascent>height||value.line_gap_millipoints!==0||center<0||!Number.isInteger(center)||source.lines.length*step>bottom-top||!Array.isArray(value.lines)||value.lines.length!==source.lines.length)throw new TypeError('Invalid centered line metrics')
 let end=0
 for(let i=0;i<value.lines.length;i++){
  const line=value.lines[i],s=source.lines[i]!;keys(line,['ordinal','start_utf16','end_utf16','baseline_millipoints','path_start','path_end'])
  if(line.ordinal!==i||line.start_utf16!==s.start_utf16||line.end_utf16!==s.end_utf16||line.baseline_millipoints!==top+center+ascent+i*step||line.path_start!==end||!Number.isSafeInteger(line.path_end)||Number(line.path_end)<end||(Number(line.path_end)===end&&/[^ ]/.test(text.slice(s.start_utf16,s.end_utf16)))||Number(line.path_end)>pathCount)throw new TypeError('Invalid line paint coverage')
  end=Number(line.path_end)
 }
 if(end!==pathCount)throw new TypeError('Incomplete line path coverage')
 return value as unknown as NativeTextboxLinePaintV1
}
