/**
 * ECMA-376 Part 1 §18.3.1.46 header and footer format codes, resolved into the
 * left, centre and right sections Excel prints.
 *
 * The codes this reads are the ones whose printed result is fully determined by
 * the worksheet and the pagination that is already known: section switches, the
 * font and size selectors, bold and italic, the worksheet name and the page
 * number and count. Everything else — a picture, a file path, the print date or
 * time, a colour or an underline — either needs a fact no preview holds or paints
 * something this tier cannot draw, so a code outside the list refuses the WHOLE
 * string rather than dropping the part it does not know. A header that silently
 * loses its date is worse than one that is honestly not painted.
 */
export type NativeSheetHeaderFooterAlignV1='left'|'center'|'right'
export interface NativeSheetHeaderFooterRunV1 {
 readonly text:string
 /** Authored face, or undefined for the workbook's own Normal font. */
 readonly font_name?:string
 readonly font_size_points:number
 readonly bold:boolean
 readonly italic:boolean
}
export interface NativeSheetHeaderFooterSectionV1 {
 readonly align:NativeSheetHeaderFooterAlignV1
 readonly runs:readonly NativeSheetHeaderFooterRunV1[]
}
/** The facts `&A`, `&P` and `&N` stand for, supplied by the paginating tier. */
export interface NativeSheetHeaderFooterFactsV1 {
 readonly sheet_name:string
 readonly page_number:number
 readonly page_count:number
 readonly default_font_size_points:number
}
const ALIGN:Record<string,NativeSheetHeaderFooterAlignV1>={L:'left',C:'center',R:'right'}
/** Excel's own header font styles; the face name may be `-` to keep the current one. */
const STYLES:Record<string,{bold:boolean;italic:boolean}>={
 'regular':{bold:false,italic:false},'bold':{bold:true,italic:false},
 'italic':{bold:false,italic:true},'bolditalic':{bold:true,italic:true},
 'bold italic':{bold:true,italic:true},'italic bold':{bold:true,italic:true},
}
/**
 * The authored string as printed sections, or undefined when any part of it is a
 * code this tier does not resolve. Sections keep authored order within each
 * alignment; an empty one is dropped because Excel prints nothing for it.
 */
export function parseNativeSheetHeaderFooterV1(code:string,facts:NativeSheetHeaderFooterFactsV1):readonly NativeSheetHeaderFooterSectionV1[]|undefined{
 if(typeof code!=='string'||code.length>1024)return undefined
 if(!Number.isFinite(facts.default_font_size_points)||facts.default_font_size_points<1||facts.default_font_size_points>409)return undefined
 const sections=new Map<NativeSheetHeaderFooterAlignV1,NativeSheetHeaderFooterRunV1[]>([['left',[]],['center',[]],['right',[]]])
 let align:NativeSheetHeaderFooterAlignV1='left',text=''
 let fontName:string|undefined,size=facts.default_font_size_points,bold=false,italic=false
 const flush=()=>{
  if(text==='')return
  sections.get(align)!.push(Object.freeze({text,...(fontName!==undefined?{font_name:fontName}:{}),font_size_points:size,bold,italic}))
  text=''
 }
 for(let i=0;i<code.length;i++){
  const ch=code[i]!
  if(ch!=='&'){text+=ch;continue}
  const next=code[i+1]
  if(next===undefined)return undefined
  i++
  if(next==='&'){text+='&';continue}
  if(ALIGN[next]){flush();align=ALIGN[next]!;continue}
  if(next==='A'){text+=facts.sheet_name;continue}
  if(next==='P'){text+=String(facts.page_number);continue}
  if(next==='N'){text+=String(facts.page_count);continue}
  if(next==='B'){flush();bold=!bold;continue}
  if(next==='I'){flush();italic=!italic;continue}
  if(next==='"'){
   const end=code.indexOf('"',i+1)
   if(end<0)return undefined
   const body=code.slice(i+1,end)
   i=end
   // `&"face,style"`: the style is everything after the LAST comma, because a
   // face name may itself contain one. An unknown style is not a face fact.
   const comma=body.lastIndexOf(',')
   if(comma<0)return undefined
   const style=STYLES[body.slice(comma+1).trim().toLowerCase().replace(/\s+/g,' ')]
   if(!style)return undefined
   const face=body.slice(0,comma)
   if(face.length>128)return undefined
   flush()
   if(face!=='-')fontName=face===''?undefined:face
   bold=style.bold;italic=style.italic
   continue
  }
  if(next>='0'&&next<='9'){
   let digits=next
   while(digits.length<3&&code[i+1]!==undefined&&code[i+1]!>='0'&&code[i+1]!<='9'){digits+=code[i+1]!;i++}
   const points=Number(digits)
   if(!Number.isInteger(points)||points<1||points>409)return undefined
   flush();size=points
   continue
  }
  return undefined
 }
 flush()
 const out:NativeSheetHeaderFooterSectionV1[]=[]
 for(const key of ['left','center','right'] as const){
  const runs=sections.get(key)!
  if(runs.length)out.push(Object.freeze({align:key,runs:Object.freeze(runs) as readonly NativeSheetHeaderFooterRunV1[]}))
 }
 return out.length?Object.freeze(out):undefined
}
