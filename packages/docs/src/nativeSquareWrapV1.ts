import type {NativeDocxDocumentV1} from './nativeContract.js'
import type {NativeDocxResolvedLayoutInputV1} from './nativeResolvedLayout.js'
import type {NativeDocxShapedLinesV1,NativeDocxLineIntervalPlanV1} from './nativeShapingLines.js'
import type {NativeDocxPaginatedLayoutV1} from './nativePaginationV1.js'
import {qualifyNativeDocxInlineImageV1} from './nativeImagePagePaintV1.js'

export function hasNativeSquareWrapV1(document:NativeDocxDocumentV1):boolean {
 return document.body.blocks.some(block=>block.paragraph?.runs.some(run=>run.drawing?.placement==='floating'&&run.drawing.wrap==='square'))
}

/** One horizontal interval only. No guesses for tight/through, middle islands,
 * or lines requiring a vertical jump. Recomputed from source on every pass. */
export function deriveNativeSquareWrapPlanV1(document:NativeDocxDocumentV1,resolved:NativeDocxResolvedLayoutInputV1,shaped:NativeDocxShapedLinesV1,layout:NativeDocxPaginatedLayoutV1,verify=false):NativeDocxLineIntervalPlanV1 {
 if(!hasNativeSquareWrapV1(document)){
  if(verify&&shaped.paragraphs.some(p=>p.lines.some(line=>line.exclusion_start_millipoints!==undefined)))throw new Error('Shaped exclusion origin has no square-wrap source')
  return {}
 }
 if(layout.status!=='paginated')throw new Error('Square wrapping requires complete provisional pagination')
 if(document.body.blocks.some(block=>block.table)||document.sections.some(section=>section.page.columns!==1)||layout.pages.some(page=>(page.note_stories?.length??0)>0))throw new Error('Square wrapping requires single-column body paragraphs without tables or notes')
 const paragraphs=new Map(shaped.paragraphs.map(p=>[p.paragraph_id,p])),properties=new Map(resolved.paragraphs.map(p=>[p.paragraph_id,p]))
 const owners=new Map<string,string>()
 let count=0
 for(const page of layout.pages)for(const placed of page.lines){
  if(++count>4096)throw new Error('Square wrapping exceeds 4096 placed lines')
  const line=paragraphs.get(placed.paragraph_id)?.lines[placed.source_line_ordinal]
  if(!line||line.id!==placed.line_id)throw new Error('Square wrapping lost source line identity')
  for(const fragment of line.fragments)if(fragment.source_kind==='image'){if(owners.has(fragment.source_id))throw new Error('Floating anchor was placed more than once');owners.set(fragment.source_id,page.id)}
 }
 const images=new Map<string,Array<{x:number;y:number;width:number;height:number}>>()
 for(const block of document.body.blocks)for(const run of block.paragraph?.runs??[]){
  if(run.drawing?.placement!=='floating'||run.drawing.wrap!=='square')continue
  const qualified=qualifyNativeDocxInlineImageV1(document,run.id,run.drawing),page=layout.pages.find(p=>p.id===owners.get(run.id))
  if(!qualified.ok||!qualified.value.floating||!page)throw new Error('Square image lacks a qualified source-bound anchor page')
  const image=qualified.value,f=image.floating!,rect={x:f.x_millipoints,y:f.y_millipoints,width:image.width_millipoints,height:image.height_millipoints},body=page.body_box
  if(rect.x+rect.width>page.width_millipoints||rect.y+rect.height>page.height_millipoints)throw new Error('Square image exceeds its anchor page')
  if(rect.x>body.x_millipoints&&rect.x+rect.width<body.x_millipoints+body.width_millipoints)throw new Error('Square image creates two text intervals; middle-image wrapping remains unsupported')
  const onPage=images.get(page.id)??[];onPage.push(rect);images.set(page.id,onPage)
 }
 const plan:Record<string,Array<{start_millipoints:number;width_millipoints:number}>>=Object.create(null)
 for(const page of layout.pages)for(const placed of page.lines){
  const paragraph=paragraphs.get(placed.paragraph_id)!,line=paragraph.lines[placed.source_line_ordinal]!,resolvedParagraph=properties.get(placed.paragraph_id)!,p=resolvedParagraph.properties,body=page.body_box
  if(p.bidi||!['left','start'].includes(p.alignment??'start')||resolvedParagraph.numbering)throw new Error('Square wrapping currently requires left-aligned LTR paragraphs without numbering')
  if(line.line_height_millipoints<line.ascent_millipoints-line.descent_millipoints)throw new Error('Square wrapping refuses clipped line metrics')
  let left=body.x_millipoints,right=left+body.width_millipoints
  for(const rect of images.get(page.id)??[]){
   if(placed.y_millipoints>=rect.y+rect.height||placed.y_millipoints+placed.height_millipoints<=rect.y||rect.x>=right||rect.x+rect.width<=left)continue
   if(rect.x<=left)left=Math.max(left,rect.x+rect.width)
   else if(rect.x+rect.width>=right)right=Math.min(right,rect.x)
   else throw new Error('Overlapping square images leave multiple text intervals')
  }
  if(right<=left)throw new Error('Square image fully blocks a text line; vertical displacement is unsupported')
  const interval={start_millipoints:left-body.x_millipoints,width_millipoints:right-left}
  ;(plan[placed.paragraph_id]??=[])[placed.source_line_ordinal]=interval
  if(verify){
   const start=(p.indent_start_twips??p.indent_left_twips??0)*50+(placed.source_line_ordinal===0?(p.first_line_twips??-(p.hanging_twips??0))*50:0),end=(p.indent_end_twips??p.indent_right_twips??0)*50
   const expectedLeft=Math.max(start,interval.start_millipoints),expectedRight=Math.min(body.width_millipoints-end,interval.start_millipoints+interval.width_millipoints)
   if(expectedRight<=expectedLeft||line.inline_offset_millipoints!==expectedLeft||line.available_width_millipoints!==expectedRight-expectedLeft||line.advance_inline_millipoints>line.available_width_millipoints||placed.x_millipoints!==body.x_millipoints+expectedLeft||line.exclusion_start_millipoints!==(expectedLeft!==start?expectedLeft:undefined))throw new Error('Final square-wrap line does not match the source-derived exclusion interval')
  }
 }
 return plan
}
