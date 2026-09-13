import type {NativePptxTableGeometryPreview} from '@injoffice/pptx-native'

/** A source arrangement inspector. Host CSS text is deliberately separate from
 * the native slide renderer and never claims authored table formatting. */
export function NativePptxTableGeometry({preview,readingId}:{preview:NativePptxTableGeometryPreview;readingId:string}) {
 return <section aria-label="Approximate source-positioned table preview" className="pptx-table-geometry">
  <p role="status">Approximate table arrangement. Stored table and cell rectangles are positioned at 96 CSS pixels per inch. This is a table-content bounding box, not a full slide.</p>
  <p className="ds-muted">Host policy: sans-serif 12 pt, 20 px line height, 2 px inset, top-left text, wrapping and clipping to each cell. {preview.paintPolicy?'Qualified one-cell tables use source no-fill and equal solid borders as a centered rectangle with round joins. Other borders remain inspection guides. Authored fonts, alignment and spacing are omitted.':'Borders below are inspection guides. Authored fonts, styles, borders, fills, alignment and spacing are omitted.'} Text may be clipped; use the full text links below.</p>
  {preview.slides.map(slide=><section key={slide.slideId} aria-label={`Table arrangement on slide ${slide.slideIndex+1}`}>
   <h4>Slide {slide.slideIndex+1} — stored table arrangement</h4>
   <div className="pptx-table-geometry-scroll" tabIndex={0} role="region" aria-label={`Scrollable table arrangement on slide ${slide.slideIndex+1}`}>
    <div className="pptx-table-geometry-plane" style={{width:slide.width+16,height:slide.height+16}}>
     {slide.tables.map(table=><div key={table.objectId} data-source-paint={table.paint?table.paint.border?'solid-border':'no-border':'omitted'} className="pptx-table-geometry-frame" data-table-object={table.objectId} style={{left:table.rect.x+8,top:table.rect.y+8,width:table.rect.width,height:table.rect.height}}>
      {table.paint?.border&&<svg aria-label={`Source border for table ${table.tableIndex+1}`} className="pptx-table-source-border" width={table.rect.width} height={table.rect.height} viewBox={`0 0 ${table.rect.width} ${table.rect.height}`}><rect x="0" y="0" width={table.rect.width} height={table.rect.height} fill="none" stroke={`#${table.paint.border.color}`} strokeWidth={table.paint.border.width_emu/9525} strokeLinejoin="round"/></svg>}
      {table.cells.map(cell=><div key={`${cell.row}:${cell.column}`} className={`pptx-table-geometry-cell${table.paint?' pptx-table-qualified-paint':''}`} data-geometry-cell={`${table.tableIndex+1}:${cell.row+1}:${cell.column+1}`} style={{left:cell.rect.x,top:cell.rect.y,width:cell.rect.width,height:cell.rect.height,fontSize:preview.fontSize,lineHeight:`${preview.lineHeight}px`,padding:preview.inset}}>
       {cell.paragraphs.map((text,index)=><p key={index} dir="auto">{text||'\u00a0'}</p>)}
      </div>)}
     </div>)}
    </div>
   </div>
   <p className="ds-muted">Source origin: x {slide.sourceBounds.x}, y {slide.sourceBounds.y} EMU. Other slide content is omitted.</p>
   <nav aria-label={`Full table text on slide ${slide.slideIndex+1}`} className="pptx-table-geometry-links">{slide.tables.map(table=><a key={table.objectId} href={`#${readingId}-${table.tableIndex}`}>Full text: table {table.tableIndex+1}</a>)}</nav>
  </section>)}
  {preview.paintPolicy&&<details open><summary>Source paint coverage</summary><p>{preview.slides.reduce((n,s)=>n+s.tables.filter(t=>t.paint).length,0)} tables have qualified source paint. Font metrics and text placement remain the declared host approximation. This is not PowerPoint rendering.</p>{!!preview.paintOmissions?.length&&<ul>{preview.paintOmissions.map(o=><li key={`${o.slideId}:${o.tableIndex}`}>Table {o.tableIndex+1}: {o.reason}</li>)}</ul>}</details>}
  {!preview.slides.length&&<p>No table arrangement fits this preview profile.</p>}
  {!!preview.omissions.length&&<details open><summary>Arrangement omissions ({preview.omissions.length})</summary><ul>{preview.omissions.map((o,i)=><li key={i}>Slide {o.slide_id}: {o.reason}</li>)}</ul></details>}
 </section>
}
