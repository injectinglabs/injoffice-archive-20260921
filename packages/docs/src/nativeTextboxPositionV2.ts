import type {TextboxAnchorContext} from './nativeTextboxAnchorLineV2.js'
import type {NativeDocxDocumentV1} from './nativeContract.js'
import type {NativeDocxPaintPageV1} from './nativePagePaintV1.js'
import type {NativeDocxTextboxGeometryItemV1,NativeDocxTextboxShapePaintV1} from './nativeTextboxGeometryPreviewV1.js'

/** Margin origins use the authored section, never the footnote-reserved body box.
 * Later source anchors supply their attested shaped insertion context.
 * Pre-text anchors retain the first-line origin. Page edges clip strokes. */
export function resolveTextboxPosition(document:NativeDocxDocumentV1,item:NativeDocxTextboxGeometryItemV1,page:NativeDocxPaintPageV1,paint:NativeDocxTextboxShapePaintV1,context?:TextboxAnchorContext):{x:number;y:number}{
 const p=item.page_anchor!
 if(p.policy==='page-offset-no-wrap-v1')return {x:p.x_emu*10/127,y:p.y_emu*10/127}
 const lines=context?[context.line]:page.lines.filter(l=>l.region==='body'&&l.paragraph_id===item.owner.paragraph_id&&l.source_line_ordinal===0),line=lines[0]
 const section=document.sections.find(s=>s.id===line?.section_id)
 if(lines.length!==1||!line||!section||page.width_millipoints!==section.page.width_twips*50||page.height_millipoints!==section.page.height_twips*50)throw new TypeError('Textbox position requires its source section and anchor line')
 // ECMA-376 17.6.19: a right binding gutter comes out of the right margin instead of the left.
 const m=section.page.margins,rtlGutter=section.page.rtl_gutter===true
 const left=(m.left_twips+(rtlGutter?0:m.gutter_twips))*50,top=m.top_twips*50,right=page.width_millipoints-(m.right_twips+(rtlGutter?m.gutter_twips:0))*50,bottom=page.height_millipoints-m.bottom_twips*50
 const odd=page.ordinal%2===0
 const horizontal=p.horizontal_relative==='insideMargin'?(odd?'leftMargin':'rightMargin'):p.horizontal_relative==='outsideMargin'?(odd?'rightMargin':'leftMargin'):p.horizontal_relative
 const vertical=p.vertical_relative==='insideMargin'?(odd?'topMargin':'bottomMargin'):p.vertical_relative==='outsideMargin'?(odd?'bottomMargin':'topMargin'):p.vertical_relative
 const hAlign=p.horizontal_align==='inside'?(odd?'left':'right'):p.horizontal_align==='outside'?(odd?'right':'left'):p.horizontal_align
 const vAlign=p.vertical_align==='inside'?(odd?'top':'bottom'):p.vertical_align==='outside'?(odd?'bottom':'top'):p.vertical_align
 const column=page.columns.find(c=>c.id===line.column_id&&c.section_id===section.id)
 let x=0,y=0,width=page.width_millipoints,height=page.height_millipoints
 if(horizontal==='leftMargin'){x=0;width=left}
 if(horizontal==='rightMargin'){x=right;width=page.width_millipoints-right}
 if(vertical==='topMargin'){y=0;height=top}
 if(vertical==='bottomMargin'){y=bottom;height=page.height_millipoints-bottom}
 if(horizontal==='margin'){x=left;width=right-left}
 if(horizontal==='column'){
  if(!column)throw new TypeError('Textbox position requires its source column')
  x=column.x_millipoints;width=column.width_millipoints
 }
 if(horizontal==='character'){x=context?.character_x??line.x_millipoints;width=0}
 if(vertical==='margin'){y=top;height=bottom-top}
 if(vertical==='paragraph'||vertical==='line'){y=vertical==='paragraph'?(context?.paragraph_y??line.y_millipoints):line.y_millipoints;height=0}
 x+=hAlign==='center'?(width-paint.width_millipoints)/2:hAlign==='right'?width-paint.width_millipoints:p.x_emu*10/127
 y+=vAlign==='center'?(height-paint.height_millipoints)/2:vAlign==='bottom'?height-paint.height_millipoints:p.y_emu*10/127
 if(!Number.isSafeInteger(x)||!Number.isSafeInteger(y))throw new TypeError('Textbox position exceeds exact coordinate precision')
 return {x,y}
}
