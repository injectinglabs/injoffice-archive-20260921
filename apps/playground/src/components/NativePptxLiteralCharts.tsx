import {useState} from 'react'
import type {NativeElement,NativePptxDeck} from '@injoffice/pptx-native'
import {createNativeLiteralPiePaths} from '@injoffice/pptx-render'

export function NativePptxLiteralCharts({deck}:{deck:NativePptxDeck}){
 const [enabled,setEnabled]=useState(false)
 const charts:Extract<NativeElement,{kind:'chart'}>[]=[]
 const visit=(elements:readonly NativeElement[])=>{for(const e of elements){if(e.kind==='chart')charts.push(e);if(e.kind==='group')visit(e.children)}}
 for(const slide of deck.slides)visit(slide.elements)
 if(!charts.length)return null
 return <section aria-label="Source literal chart preview">
  <label><input type="checkbox" checked={enabled} onChange={event=>setEnabled(event.target.checked)}/> Preview supported source literal pie charts</label>
  <p>Read-only vectors from explicit source values and colors. Each chart is shown separately using a centered circle and polygon arcs. This preview does not reproduce PowerPoint layout. Formula caches, labels, legends and other chart types are unsupported.</p>
  {enabled&&charts.map(chart=>{
   const pie=chart.chart.literalPie
   if(!pie)return <p key={chart.id}>{chart.name??chart.id}: outside the source literal pie profile; original chart preserved.</p>
   const paths=createNativeLiteralPiePaths(pie,chart.transform.cx,chart.transform.cy)
   return <figure key={chart.id}><figcaption>{chart.name??chart.id} · source literal values · preserved, read-only</figcaption>
    <svg role="img" aria-label={`${chart.name??'Pie chart'}: ${pie.values.join(', ')}`} viewBox={`0 0 ${chart.transform.cx} ${chart.transform.cy}`} style={{width:320,maxWidth:'100%',height:220}}>
     {paths.map((slice,i)=><path key={i} fill={slice.color} d={slice.path.map(p=>p.kind==='moveTo'?`M ${p.x} ${p.y}`:p.kind==='lineTo'?`L ${p.x} ${p.y}`:p.kind==='close'?'Z':'').join(' ')}/>)}
    </svg><p>Values: {pie.values.join(', ')}. Angle: {pie.firstSliceAngle}° clockwise from up.</p>
   </figure>
  })}
 </section>
}
