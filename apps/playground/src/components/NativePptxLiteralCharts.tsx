import {useState} from 'react'
import type {NativeElement,NativePptxDeck} from '@injoffice/pptx-native'
import {createNativeLiteralPiePaths,createNativeLiteralDoughnutPaths,createNativeLiteralBarPaths} from '@injoffice/pptx-render'

export function NativePptxLiteralCharts({deck}:{deck:NativePptxDeck}){
 const [enabled,setEnabled]=useState(false)
 const charts:Extract<NativeElement,{kind:'chart'}>[]=[]
 const visit=(elements:readonly NativeElement[])=>{for(const e of elements){if(e.kind==='chart')charts.push(e);if(e.kind==='group')visit(e.children)}}
 for(const slide of deck.slides)visit(slide.elements)
 if(!charts.length)return null
 return <section aria-label="Source literal chart preview">
  <label><input type="checkbox" checked={enabled} onChange={event=>setEnabled(event.target.checked)}/> Preview supported source literal charts</label>
  <p>Read-only vectors from explicit source values and colors. Each chart is shown separately with host frame fitting. Circular charts use polygon arcs; clustered bars use explicit linear scales. This preview does not reproduce PowerPoint plot layout. Formula caches, source labels and legends are unsupported.</p>
  {enabled&&charts.map(chart=>{
   const bar=chart.chart.literalBar
   if(bar){
    const extent=bar.barDirection==='column'?chart.transform.cx:chart.transform.cy
    if(BigInt(extent)*100n<BigInt(bar.categories.length*(100*bar.series.length+bar.gapWidth)))return <p key={chart.id}>{chart.name??chart.id}: frame too small to retain distinct bars; original chart preserved.</p>
    const vectors=createNativeLiteralBarPaths(bar,chart.transform.cx,chart.transform.cy)
    return <figure key={chart.id}><figcaption>{chart.name??chart.id} · source literal {bar.barDirection==='column'?'columns':'bars'} · preserved, read-only</figcaption>
     <svg role="img" aria-label={`${chart.name??'Bar chart'}: ${bar.categories.length} categories, ${bar.series.length} series. Values are listed below.`} viewBox={`0 0 ${chart.transform.cx} ${chart.transform.cy}`} style={{width:400,maxWidth:'100%',height:260}}>
      {vectors.map((vector,i)=><path key={i} fill={vector.color??'none'} stroke={vector.stroke?.color} strokeWidth={vector.stroke?.widthEmu} strokeLinecap="butt" d={vector.path.map(p=>p.kind==='moveTo'?`M ${p.x} ${p.y}`:p.kind==='lineTo'?`L ${p.x} ${p.y}`:p.kind==='close'?'Z':'').join(' ')}/>)}
     </svg><p>Explicit value scale: {bar.valueAxis.min} to {bar.valueAxis.max}. Gap: {bar.gapWidth}% of one bar width. Values outside the scale are clipped; zero and sub-EMU heights may have no visible area.</p>
     <table><caption>Source data (host table, separate from slide labels)</caption><thead><tr><th>Category</th>{bar.series.map(series=><th key={series.index}>{series.title??`Series ${series.index}`}</th>)}</tr></thead><tbody>{bar.categories.map((category,i)=><tr key={i}><th>{category}</th>{bar.series.map(series=><td key={series.index}>{series.values[i]}</td>)}</tr>)}</tbody></table>
    </figure>
   }
   const pie=chart.chart.literalPie??chart.chart.literalDoughnut
   if(!pie)return <p key={chart.id}>{chart.name??chart.id}: outside the supported source literal chart profiles; original chart preserved.</p>
   if(pie.profile==='literal-doughnut-v1' && Math.min(chart.transform.cx,chart.transform.cy)*Math.min(pie.holeSize,100-pie.holeSize)/200<1)return <p key={chart.id}>{chart.name??chart.id}: frame too small to retain the doughnut hole; original chart preserved.</p>
   const paths=pie.profile==='literal-doughnut-v1'?createNativeLiteralDoughnutPaths(pie,chart.transform.cx,chart.transform.cy):createNativeLiteralPiePaths(pie,chart.transform.cx,chart.transform.cy)
   return <figure key={chart.id}><figcaption>{chart.name??chart.id} · source literal values · preserved, read-only</figcaption>
    <svg role="img" aria-label={`${chart.name??(pie.profile==='literal-doughnut-v1'?'Doughnut chart':'Pie chart')}: ${pie.values.join(', ')}`} viewBox={`0 0 ${chart.transform.cx} ${chart.transform.cy}`} style={{width:320,maxWidth:'100%',height:220}}>
     {paths.map((slice,i)=><path key={i} fill={slice.color} d={slice.path.map(p=>p.kind==='moveTo'?`M ${p.x} ${p.y}`:p.kind==='lineTo'?`L ${p.x} ${p.y}`:p.kind==='close'?'Z':'').join(' ')}/>)}
    </svg><p>Values: {pie.values.join(', ')}. Angle: {pie.firstSliceAngle}° clockwise from up.{pie.profile==='literal-doughnut-v1'?` Hole: ${pie.holeSize}%.`:''}</p>
   </figure>
  })}
 </section>
}
