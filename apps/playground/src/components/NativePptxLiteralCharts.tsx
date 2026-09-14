import {useId,useState} from 'react'
import type {NativeElement,NativePptxDeck} from '@injoffice/pptx-native'
import {createNativeLiteralAreaPaths,createNativeLiteralPiePaths,createNativeLiteralDoughnutPaths,createNativeLiteralBarPaths,createNativeLiteralLinePaths,createNativeLiteralScatterPaths} from '@injoffice/pptx-render'

export function NativePptxLiteralCharts({deck}:{deck:NativePptxDeck}){
 const [enabled,setEnabled]=useState(false)
 const clipPrefix=useId()
 const charts:Extract<NativeElement,{kind:'chart'}>[]=[]
 const visit=(elements:readonly NativeElement[])=>{for(const e of elements){if(e.kind==='chart')charts.push(e);if(e.kind==='group')visit(e.children)}}
 for(const slide of deck.slides)visit(slide.elements)
 if(!charts.length)return null
 return <section aria-label="Source literal chart preview">
  <label><input type="checkbox" checked={enabled} onChange={event=>setEnabled(event.target.checked)}/> Preview supported source literal charts</label>
  <p>Read-only vectors from explicit source values and colors. Each chart is shown separately with host frame fitting. Circular charts use polygon arcs; Cartesian charts use explicit linear scales and clipped data geometry. This preview does not reproduce PowerPoint plot layout. Formula caches, source labels and legends are unsupported.</p>
  {enabled&&charts.map((chart,chartIndex)=>{
   const area=chart.chart.literalArea
   const connected=chart.chart.literalConnected
   if(area?.xAxis.labels||area?.yAxis.labels||connected?.xAxis.labels||connected?.yAxis.labels||chart.chart.literalBar?.categoryAxis.labels||chart.chart.literalBar?.valueAxis.labels)return <p key={chart.id}>{chart.name??chart.id}: source axis labels require the supplied-font slide preview. Original chart preserved.</p>
   if(area){
    const vectors=createNativeLiteralAreaPaths(area,chart.transform.cx,chart.transform.cy)
    return <figure key={chart.id}><figcaption>{chart.name??chart.id} · source literal {area.grouping} area · preserved, read-only</figcaption>
     <svg role="img" aria-label={`${chart.name??'Area chart'}: ${area.series.length} series. Source values are listed below.`} viewBox={`0 0 ${chart.transform.cx} ${chart.transform.cy}`} style={{width:400,maxWidth:'100%',height:260,overflow:'hidden'}}>
      <defs><clipPath id={`${clipPrefix}-${chartIndex}`}><rect x={0} y={0} width={chart.transform.cx} height={chart.transform.cy}/></clipPath></defs>
      <g clipPath={`url(#${clipPrefix}-${chartIndex})`}>{vectors.map((vector,i)=><path key={i} fill={vector.color??'none'} stroke={vector.stroke?.color} strokeWidth={vector.stroke?.widthEmu} strokeLinecap="butt" d={vector.path.map(p=>p.kind==='moveTo'?`M ${p.x} ${p.y}`:p.kind==='lineTo'?`L ${p.x} ${p.y}`:p.kind==='close'?'Z':'').join(' ')}/>)}</g>
     </svg><p>Explicit value scale: {area.yAxis.min} to {area.yAxis.max}. Each series has one compound fill. Standard overlaps paint in authored series order as a host preview policy; PowerPoint overlap order is not established. Negative stacked values are unsupported. Zero-total percentages and zero-area bands have no fill.</p>
     <table><caption>Source data (host table, separate from slide labels)</caption><thead><tr><th>Category</th>{area.series.map(series=><th key={series.index}>{series.title??`Series ${series.index}`}</th>)}</tr></thead><tbody>{area.categories.map((category,i)=><tr key={i}><th>{category}</th>{area.series.map(series=><td key={series.index}>{series.values[i]}</td>)}</tr>)}</tbody></table>
    </figure>
   }
   if(connected){
    const scatter=connected.profile==='literal-scatter-v1'
    const vectors=(scatter?createNativeLiteralScatterPaths:createNativeLiteralLinePaths)(connected,chart.transform.cx,chart.transform.cy)
    return <figure key={chart.id}><figcaption>{chart.name??chart.id} · source literal {scatter?'XY scatter':'category line'} · preserved, read-only</figcaption>
     <svg role="img" aria-label={`${chart.name??'Connected chart'}: ${connected.series.length} series. Source values are listed below.`} viewBox={`0 0 ${chart.transform.cx} ${chart.transform.cy}`} style={{width:400,maxWidth:'100%',height:260,overflow:'hidden'}}>
      <defs><clipPath id={`${clipPrefix}-${chartIndex}`}><rect x={0} y={0} width={chart.transform.cx} height={chart.transform.cy}/></clipPath></defs>
      <g clipPath={`url(#${clipPrefix}-${chartIndex})`}>{vectors.filter(vector=>vector.path.length>0).map((vector,i)=><path key={i} fill="none" stroke={vector.stroke.color} strokeWidth={vector.stroke.widthEmu} strokeLinecap="butt" strokeLinejoin="round" d={vector.path.map(p=>p.kind==='moveTo'?`M ${p.x} ${p.y}`:p.kind==='lineTo'?`L ${p.x} ${p.y}`:'').join(' ')}/>)}</g>
     </svg><p>Straight segments follow source point order. Explicit Y scale: {connected.yAxis.min} to {connected.yAxis.max}.{scatter?` X scale: ${connected.xAxis.min} to ${connected.xAxis.max}.`:''} Segments and stroke envelopes are clipped to the plot frame. Singleton or fully clipped series have no visible line.</p>
     <table><caption>Source data (host table, separate from slide labels)</caption><thead><tr><th>Series</th><th>Point</th><th>{scatter?'X':'Category'}</th><th>Y</th></tr></thead><tbody>{connected.series.flatMap(series=>series.values.map((value,i)=><tr key={`${series.index}:${i}`}><th>{series.title??`Series ${series.index}`}</th><td>{i}</td><td>{scatter?series.xValues![i]:connected.categories[i]}</td><td>{value}</td></tr>))}</tbody></table>
    </figure>
   }
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
