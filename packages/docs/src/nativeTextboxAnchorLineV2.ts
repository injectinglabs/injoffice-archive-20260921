import type {NativeDocxDocumentV1} from './nativeContract.js'
import type {NativeDocxPagePaintRequestV1,NativeDocxPagePaintSuccessV1,NativeDocxPaintLineV1} from './nativePagePaintV1.js'
import type {NativeDocxTextboxGeometryItemV1} from './nativeTextboxGeometryPreviewV1.js'

export function textboxHasPrecedingRuns(document:NativeDocxDocumentV1,item:NativeDocxTextboxGeometryItemV1):boolean {
 return document.body.blocks.find(b=>b.id===item.owner.paragraph_id)?.paragraph?.runs.some(r=>r.anchor.end_byte<item.page_anchor!.source_anchor.start_byte)??false
}

/** Locate the source insertion boundary using shaped advances, including spaces,
 * tabs and bidi fragments. Drawing anchors consume no inline advance. */
export function textboxAnchorLine(document:NativeDocxDocumentV1,item:NativeDocxTextboxGeometryItemV1,body:NativeDocxPagePaintSuccessV1,request?:NativeDocxPagePaintRequestV1){
 const paragraph=document.body.blocks.find(b=>b.id===item.owner.paragraph_id)?.paragraph
 if(!paragraph)throw new TypeError('Textbox anchor paragraph is missing')
 const placed=body.pages.flatMap(page=>page.lines.filter(l=>l.region==='body'&&l.paragraph_id===paragraph.id).map(line=>({page,line}))).sort((a,b)=>a.line.source_line_ordinal-b.line.source_line_ordinal)
 if(placed.some((e,i)=>i>0&&e.line.source_line_ordinal<=placed[i-1]!.line.source_line_ordinal))throw new TypeError('Textbox anchor lines are ambiguous')
 if(!placed.length||placed[0]!.line.source_line_ordinal!==0)throw new TypeError('Textbox anchor paragraph has no first line')
 const preceding=paragraph.runs.filter(r=>r.anchor.end_byte<item.page_anchor!.source_anchor.start_byte)
 if(!preceding.length)return {...placed[0]!,character_x:placed[0]!.line.x_millipoints,paragraph_y:placed[0]!.line.y_millipoints}
 if(!request)throw new TypeError('Textbox anchor requires shaped source layout')
 const shaped=(entry:typeof placed[number])=>{
  const p=(request.page_field_variants?.find(v=>v.page_id===entry.page.id)?.shaped_lines??request.pagination_request.shaped_lines).paragraphs.find(p=>p.paragraph_id===paragraph.id)
  const line=p?.lines[entry.line.source_line_ordinal]
  if(!line||line.id!==entry.line.line_id)throw new TypeError('Textbox anchor line does not match source shaping')
  const sourcePage=request.paginated_layout.pages.find(p=>p.id===entry.page.id),sourceLine=sourcePage?.lines.find(l=>l.id===entry.line.placed_line_id)
  if(!sourceLine||sourceLine.line_id!==line.id||sourceLine.x_millipoints!==entry.line.x_millipoints||sourceLine.y_millipoints!==entry.line.y_millipoints||sourceLine.width_millipoints!==entry.line.width_millipoints||sourceLine.height_millipoints!==entry.line.height_millipoints||line.advance_inline_millipoints!==entry.line.width_millipoints||entry.line.y_millipoints+line.ascent_millipoints!==entry.line.baseline_y_millipoints||line.fragments.reduce((n,f)=>n+f.advance_inline_millipoints,0)!==line.advance_inline_millipoints)throw new TypeError('Textbox anchor metrics do not match body paint')
  const logical=new Set<number>()
  for(const f of line.fragments){
   if(!Number.isSafeInteger(f.advance_inline_millipoints)||f.advance_inline_millipoints<0||!Number.isSafeInteger(f.logical_order)||f.logical_order<0||logical.has(f.logical_order)||!['ltr','rtl'].includes(f.direction))throw new TypeError('Invalid textbox anchor fragment metrics')
   logical.add(f.logical_order)
   if(f.source_kind==='list-marker')continue
   const run=paragraph.runs.find(r=>r.id===f.source_id)
   if(!run)throw new TypeError('Textbox anchor fragment has no source run')
   if(run.kind==='text'&&!run.page_field&&(!Number.isSafeInteger(f.start_utf16)||!Number.isSafeInteger(f.end_utf16)||f.start_utf16<0||f.end_utf16<f.start_utf16||f.end_utf16>(run.text?.length??0)||f.text!==run.text?.slice(f.start_utf16,f.end_utf16)))throw new TypeError('Textbox anchor fragment text does not match source')
  }
  return line
 }
 let selected=placed[0]!,x=selected.line.x_millipoints,found=false
 for(const run of preceding.slice().reverse()){
  if(run.kind==='text'&&run.text===''&&!run.page_field)continue
  for(const entry of placed.slice().reverse()){
   const line=shaped(entry)
   if(line.hard_break_after?.source_run_id===run.id){
    const next=placed.find(e=>e.line.source_line_ordinal===entry.line.source_line_ordinal+1)
    if(!next)throw new TypeError('Textbox anchor after break has no following line')
    const nextLine=shaped(next),first=nextLine.fragments.map((fragment,index)=>({fragment,index})).filter(f=>f.fragment.source_kind!=='list-marker').sort((a,b)=>a.fragment.logical_order-b.fragment.logical_order)[0]
    selected=next;x=next.line.x_millipoints+(first?nextLine.fragments.slice(0,first.index).reduce((n,f)=>n+f.advance_inline_millipoints,0)+(first.fragment.direction==='rtl'?first.fragment.advance_inline_millipoints:0):0);found=true;break
   }
   const fragments=line.fragments.map((fragment,index)=>({fragment,index})).filter(f=>f.fragment.source_id===run.id).sort((a,b)=>b.fragment.end_utf16-a.fragment.end_utf16||b.fragment.logical_order-a.fragment.logical_order)
   if(!fragments.length)continue
   const {fragment,index}=fragments[0]!
   x=entry.line.x_millipoints+line.fragments.slice(0,index).reduce((n,f)=>n+f.advance_inline_millipoints,0)+(fragment.direction==='ltr'?fragment.advance_inline_millipoints:0)
   selected=entry;found=true;break
  }
  if(found)break
  throw new TypeError('Textbox preceding source run has no shaped insertion boundary')
 }
 const paragraphLine=selected.page.lines.filter(l=>l.region==='body'&&l.paragraph_id===paragraph.id).sort((a,b)=>a.source_line_ordinal-b.source_line_ordinal)[0]!
 return {...selected,character_x:x,paragraph_y:paragraphLine.y_millipoints}
}
export type TextboxAnchorContext={line:NativeDocxPaintLineV1;character_x:number;paragraph_y:number}
