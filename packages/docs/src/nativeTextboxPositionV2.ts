import type {NativeDocxDocumentV1} from './nativeContract.js'
import type {NativeDocxPaintPageV1} from './nativePagePaintV1.js'
import type {NativeDocxTextboxGeometryItemV1,NativeDocxTextboxShapePaintV1} from './nativeTextboxGeometryPreviewV1.js'

/** Margin origins use the authored section, never the footnote-reserved body box.
 * Drawing anchors precede all modeled text, so paragraph/line/character origins
 * are the owner's first placed line. Alignment edges clip strokes at the page. */
export function resolveTextboxPosition(document:NativeDocxDocumentV1,item:NativeDocxTextboxGeometryItemV1,page:NativeDocxPaintPageV1,paint:NativeDocxTextboxShapePaintV1):{x:number;y:number}{
 const p=item.page_anchor!
 if(p.policy==='page-offset-no-wrap-v1')return {x:p.x_emu*10/127,y:p.y_emu*10/127}
 const lines=page.lines.filter(l=>l.region==='body'&&l.paragraph_id===item.owner.paragraph_id&&l.source_line_ordinal===0),line=lines[0]
 const section=document.sections.find(s=>s.id===line?.section_id)
 if(lines.length!==1||!line||!section||page.width_millipoints!==section.page.width_twips*50||page.height_millipoints!==section.page.height_twips*50)throw new TypeError('Textbox position requires its source section and first line')
 const m=section.page.margins,left=(m.left_twips+m.gutter_twips)*50,top=m.top_twips*50,right=page.width_millipoints-m.right_twips*50,bottom=page.height_millipoints-m.bottom_twips*50
 const column=page.columns.find(c=>c.id===line.column_id&&c.section_id===section.id)
 let x=0,y=0,width=page.width_millipoints,height=page.height_millipoints
 if(p.horizontal_relative==='margin'){x=left;width=right-left}
 if(p.horizontal_relative==='column'){
  if(!column)throw new TypeError('Textbox position requires its source column')
  x=column.x_millipoints;width=column.width_millipoints
 }
 if(p.horizontal_relative==='character'){x=line.x_millipoints;width=0}
 if(p.vertical_relative==='margin'){y=top;height=bottom-top}
 if(p.vertical_relative==='paragraph'||p.vertical_relative==='line'){y=line.y_millipoints;height=0}
 x+=p.horizontal_align==='center'?(width-paint.width_millipoints)/2:p.horizontal_align==='right'?width-paint.width_millipoints:p.x_emu*10/127
 y+=p.vertical_align==='center'?(height-paint.height_millipoints)/2:p.vertical_align==='bottom'?height-paint.height_millipoints:p.y_emu*10/127
 if(!Number.isSafeInteger(x)||!Number.isSafeInteger(y))throw new TypeError('Textbox position exceeds exact coordinate precision')
 return {x,y}
}
