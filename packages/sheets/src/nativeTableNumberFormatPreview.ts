import type {NativeWorkbookObjectsV1} from './nativeObjectsPreviewV1.js'
import {formatNativeSheetCellDisplayV2} from './nativeCellDisplayV2.js'

export type NativeTableNumberDisplayV1={text:string;warning?:string}|{warning:string;text?:never}

function contains(ref:string,row:number,column:number):boolean{
 const match=/^([A-Z]{1,3})([1-9][0-9]{0,6}):([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(ref);if(!match)return false
 const col=(s:string)=>[...s].reduce((v,c)=>v*26+c.charCodeAt(0)-64,0)-1
 const left=col(match[1]!),right=col(match[3]!),top=Number(match[2])-1,bottom=Number(match[4])-1
 return left<=right&&top<=bottom&&right<16384&&bottom<1048576&&row>=top&&row<=bottom&&column>=left&&column<=right
}

/** Read-only differential display, using saved lexical values without recalculation. */
export function nativeTableNumberFormatPreview(objects:NativeWorkbookObjectsV1,revision:string,sheetPart:string,row:number,column:number,styleID:number,kind:'number'|'date',lexical:string,date1904?:boolean):NativeTableNumberDisplayV1|undefined{
 if(objects.package_sha256!==revision||!Number.isSafeInteger(row)||!Number.isSafeInteger(column))return undefined
 const tables=objects.tables.filter(t=>t.sheet_part===sheetPart&&contains(t.ref,row,column));if(tables.length!==1)return undefined
 const regions=tables[0]!.number_formats?.filter(f=>contains(f.ref,row,column));if(!regions?.length)return undefined
 if(regions.length!==1)return {warning:'Overlapping table number formats are not resolved.'}
 const region=regions[0]!;if(!region.style_ids.includes(styleID))return undefined
 const direct=formatNativeSheetCellDisplayV2(kind,lexical,region.number_format,date1904)
 if(direct.status==='ready')return {text:direct.text}
 const accounting=kind==='number'?formatNativeAccountingTextPreview(lexical,region.number_format):undefined
 return accounting??{warning:'Unsupported table differential number format; showing the stored value.'}
}

// Display-only text fallback. Padding/fill directives remain an explicit layout
// limitation; the strict glyph-paint formatter is deliberately not broadened.
export function formatNativeAccountingTextPreview(lexical:string,format:string):NativeTableNumberDisplayV1|undefined{
 if(format.length>4096||lexical.length>32767||!/^[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?$/.test(lexical))return undefined
 const sections=[''];let quoted=false,escaped=false
 for(const char of format){if(escaped){sections[sections.length-1]+=char;escaped=false;continue}if(char==='\\'){sections[sections.length-1]+=char;escaped=true;continue}if(char==='"')quoted=!quoted;if(char==='['&&!quoted)return undefined;if(char===';'&&!quoted){if(sections.length===4)return undefined;sections.push('')}else sections[sections.length-1]+=char}
 if(quoted||escaped)return undefined
 const zero=!/[1-9]/.test(lexical.split(/[eE]/)[0]!),negative=lexical.startsWith('-')&&!zero
 const section=sections[zero&&sections.length>=3?2:negative&&sections.length>=2?1:0]!
 let prefix='',suffix='',numeric='',seen=false,padding=false
 const literal=(text:string)=>{if(seen)suffix+=text;else prefix+=text}
 for(let i=0;i<section.length;i++){
  const c=section[i]!
  if(c==='_'||c==='*'){if(i+1>=section.length)return undefined;i++;padding=true;continue}
  if(c==='\\'){if(i+1>=section.length)return undefined;literal(section[++i]!);continue}
  if(c==='"'){let text='';while(++i<section.length&&section[i]!=='"')text+=section[i];if(i===section.length||!/^[\x20-\x7e]*$/.test(text))return undefined;literal(text);continue}
  if(/[0#?,.%]/.test(c)){if(suffix)return undefined;seen=true;numeric+=c;continue}
  if(c===' '||c==='-'||c==='+'||c==='('||c===')'){literal(c);continue}
  return undefined
 }
 let text:string
 if(zero&&/^\?*$/.test(numeric)&&prefix+suffix){text=prefix+suffix;padding||=numeric.length>0}
 else {
  const amount=lexical.replace(/^[-+]/,'')
  const result=formatNativeSheetCellDisplayV2('number',amount,numeric)
  if(result.status!=='ready'||numeric==='General')return undefined
  text=`${negative&&sections.length===1?'-':''}${prefix}${result.text}${suffix}`
 }
 if(!text||text.length>32767)return undefined
 return {text,...(padding?{warning:'Accounting fill and padding alignment are not reproduced; displayed digits and literals use the saved value.'}:{})}
}
