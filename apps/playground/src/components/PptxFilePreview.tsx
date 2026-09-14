import {useEffect,useId,useState,type ReactNode} from 'react'
import type {NativeElement,NativePptxDeck} from '@injoffice/pptx-native'
import type {PaintCommand,RenderNode,RenderRect,RenderTextBodyNode} from '@injoffice/pptx-render'
import {DsButton,DsField,DsSelect} from '../design-system/primitives'
import {previewColor,previewImage,previewIssue,previewSlideSize} from '../pptxPreview'
import {compileFilePreviewGeometry,filePreviewLength as n,filePreviewPath,type FilePreviewGeometry} from '../filePreviewGeometry'

export function PptxFilePreviewVector({deck,geometry}:{deck:NativePptxDeck;geometry:FilePreviewGeometry}){
 const prefix=useId().replace(/:/g,''),matrix=(value:RenderNode['transform'])=>`matrix(${geometry.matrices.get(value)!.join(' ')})`
 const issueFor=(element:NativeElement)=>element.kind==='table'&&element.graphicFrameLayout==='source-anchored-v1'&&element.compatibility.status==='preserveOnly'?undefined:previewIssue(element,deck.assets)
 const rectProps=(r:RenderRect)=>({x:n(r.x),y:n(r.y),width:n(r.cx),height:n(r.cy)})
 function path(command:Extract<PaintCommand,{kind:'path'}>,key:string):ReactNode{
  const stroke=command.stroke,props={fill:previewColor(command.fill,'none'),stroke:previewColor(stroke?.color,'none'),strokeWidth:stroke?n(stroke.widthEmu):undefined,strokeLinecap:stroke?.cap==='flat'?'butt' as const:stroke?.cap,strokeLinejoin:stroke?.join,strokeMiterlimit:stroke?.miterLimit}
  const first=command.path[0]
  if(command.path.length===1&&first?.kind==='rect')return <rect key={key} {...rectProps(first.rect)} {...props}/>
  if(command.path.length===1&&first?.kind==='roundRect')return <rect key={key} {...rectProps(first.rect)} rx={n(first.radiusEmu)} {...props}/>
  if(command.path.length===1&&first?.kind==='ellipse')return <ellipse key={key} cx={n(first.rect.x+first.rect.cx/2)} cy={n(first.rect.y+first.rect.cy/2)} rx={n(first.rect.cx/2)} ry={n(first.rect.cy/2)} {...props}/>
  return <path key={key} d={filePreviewPath(command.path)} {...props}/>
 }
 function approximateText(original:Pick<Extract<NativeElement,{kind:'shape'|'text'}>,'id'|'paragraphs'|'textBody'>,bounds:RenderRect,body?:RenderTextBodyNode):ReactNode{
  if(!original.paragraphs.length)return null
  if(body?.status==='refused'||geometry.textIssues.has(original.id))return <text {...{x:n(bounds.x),y:n(bounds.y)+12}} fontSize={10}>Text preview unavailable</text>
  const vertical=Boolean(body?.transform),box=vertical?{x:0,y:0,cx:bounds.cy,cy:bounds.cx}:bounds
  const content=<foreignObject data-file-preview-text={original.id} {...rectProps(box)} style={{overflow:original.textBody?'visible':'hidden'}}>
   <div style={{height:'100%',boxSizing:'border-box',overflow:original.textBody?'visible':'hidden',color:'#141816',display:'flex',flexDirection:'column',justifyContent:original.textBody?.verticalAnchor==='center'?'center':original.textBody?.verticalAnchor==='bottom'?'flex-end':'flex-start'}}>
    {original.paragraphs.map((paragraph,i)=><p key={i} style={{margin:0,paddingLeft:n(paragraph.marginLeftEmu??0),textIndent:n(paragraph.indentEmu??0),flexShrink:0,lineHeight:1.2,whiteSpace:original.textBody?.wrap==='none'?'pre':'pre-wrap',overflowWrap:'normal',textAlign:paragraph.align==='center'?'center':paragraph.align==='right'?'right':'left',fontFamily:paragraph.runs[0]?.fontFamily??'Arial, sans-serif',fontSize:n((paragraph.runs[0]?.fontSizeHundredthPt??1800)*127)}}>
     {paragraph.bullet&&<span style={{color:previewColor(paragraph.runs[0]?.color,'#141816'),fontWeight:paragraph.runs[0]?.bold?700:400,fontStyle:paragraph.runs[0]?.italic?'italic':'normal'}}>{`${paragraph.bulletCharacter??'•'} `}</span>}{paragraph.runs.map((run,j)=><span key={j} style={{fontFamily:run.fontFamily??'Arial, sans-serif',fontWeight:run.bold?700:400,fontStyle:run.italic?'italic':'normal',fontSize:n((run.fontSizeHundredthPt??1800)*127),color:previewColor(run.color,'#141816')}}>{run.text}</span>)}
    </p>)}
   </div>
  </foreignObject>
  const rotated=body?.transform?<g transform={matrix(body.transform)}>{content}</g>:content
  return body?.orientationTransform?<g transform={matrix(body.orientationTransform)}>{rotated}</g>:rotated
 }
 function object(node:RenderNode,key:string):ReactNode{
  const original=geometry.originals.get(node.sourceElementId)
  if(!original)return null
  const issue=issueFor(original),clip=node.clip,id=`${prefix}-${key}`
  const placeholder=(label:string)=><><rect {...rectProps(node.bounds)} fill="#f5f6f6" stroke="#5a6560" strokeDasharray="3 2"/><text x={n(node.bounds.x)+4} y={n(node.bounds.y)+14} fontSize={10}>{label}</text></>
  let content:ReactNode
  if(node.kind==='placeholder'||issue)content=placeholder(`${original.kind} preview unavailable`)
  else if(node.kind==='group')content=node.children.map((child,i)=>object(child,`${key}-${i}`))
  else if(node.kind==='image'){
   const raster=previewImage(deck.assets.find(asset=>asset.id===node.assetId)),crop=node.crop
   content=raster?(crop?<svg {...rectProps(node.bounds)} viewBox={`${crop.left} ${crop.top} ${100000-crop.left-crop.right} ${100000-crop.top-crop.bottom}`} preserveAspectRatio="none" overflow="hidden"><image href={raster} width={100000} height={100000} preserveAspectRatio="none"/></svg>:<image href={raster} {...rectProps(node.bounds)} preserveAspectRatio="none"/>):placeholder('Image preview unavailable')
  }else if(node.kind==='shape'||node.kind==='connector'||node.kind==='text'){
   content=<>{(geometry.paths.get(node.sourceElementId)??[]).map((command,i)=>path(command,`${key}-path-${i}`))}{(original.kind==='shape'||original.kind==='text')&&approximateText(original,(node.kind==='shape'||node.kind==='text')&&node.textBody?node.textBody.bounds:node.bounds,node.kind==='shape'||node.kind==='text'?node.textBody:undefined)}</>
  }else if(node.kind==='table'&&original.kind==='table'&&original.graphicFrameLayout==='source-anchored-v1'){
   content=<>{node.cells.map((cell,i)=><rect key={`background-${i}`} data-file-preview-cell-background={`${original.id}-${i}`} {...rectProps(cell.bounds)} fill={previewColor(cell.fill?.color,'none')} stroke={previewColor(cell.border?.color,'none')} strokeWidth={cell.border?n(cell.border.widthEmu):undefined}/>)}{node.cells.map((cell,i)=>{
    const source=original.table.rows[cell.rowIndex]?.[cell.columnIndex]
    if(!source)return null
    const textSource={id:`${original.id}-cell-${i}`,paragraphs:source.paragraphs??[{runs:[{text:source.text}]}],textBody:source.textBody}
    const text=approximateText(textSource,cell.textBody?.bounds??{x:0,y:0,cx:cell.bounds.cx,cy:cell.bounds.cy},cell.textBody)
    // Browser text remains explicitly approximate. This x strip spans the
    // complete slide in cell-local coordinates; the outer slide clip supplies
    // the y boundary, without claiming native measured browser glyph bounds.
    const tableY=geometry.matrices.get(node.transform)![5]!*12700
    const strip={x:0,y:-tableY-cell.bounds.y,cx:cell.bounds.cx,cy:geometry.tree.size.cy}
    const clipId=`${id}-cell-${i}`
    return <g key={i} data-file-preview-cell={`${original.id}-${i}`} transform={`translate(${n(cell.bounds.x)} ${n(cell.bounds.y)})`}>
     {source.textBody?.horizontalOverflow==='clip'&&<defs><clipPath id={clipId} clipPathUnits="userSpaceOnUse"><rect {...rectProps(strip)}/></clipPath></defs>}
     <g clipPath={source.textBody?.horizontalOverflow==='clip'?`url(#${clipId})`:undefined}>{text}</g>
    </g>
   })}</>
  }else content=placeholder('Table preview unavailable')
  return <g key={key} data-file-preview-object={original.id} transform={matrix(node.transform)}><title>{`${original.name||original.kind}${issue?`: ${issue}`:''}`}</title>
   {clip&&<defs><clipPath id={id} clipPathUnits="userSpaceOnUse"><rect {...rectProps(clip.rect)} rx={clip.kind==='roundRect'?n(clip.radiusEmu):undefined}/></clipPath></defs>}
   <g clipPath={clip?`url(#${id})`:undefined}>{content}</g>
  </g>
 }
 return <svg data-file-preview-document={deck.documentId} viewBox={`0 0 ${n(geometry.tree.size.cx)} ${n(geometry.tree.size.cy)}`} role="img" aria-label={`Approximate preview of slide ${geometry.tree.slideIndex+1}`} style={{display:'block',width:'100%',border:'1px solid var(--ds-line)',background:previewColor(geometry.tree.background.color),overflow:'hidden'}}>{geometry.tree.nodes.map((node,i)=>object(node,String(i)))}</svg>
}

export default function PptxFilePreview({deck}:{deck:NativePptxDeck}){
 const [at,setAt]=useState(0),[result,setResult]=useState<{deck:NativePptxDeck;index:number;geometry?:FilePreviewGeometry;error?:string}>()
 useEffect(()=>setAt(0),[deck.documentId])
 const index=Math.min(at,Math.max(0,deck.slides.length-1)),slide=deck.slides[index],size=previewSlideSize(deck)
 useEffect(()=>{
  let current=true
  void compileFilePreviewGeometry(deck,index).then(geometry=>{if(current)setResult({deck,index,geometry})},error=>{if(current)setResult({deck,index,error:error instanceof Error?error.message:'Geometry preview unavailable'})})
  return()=>{current=false}
 },[deck,index])
 if(!slide||!size)return <p className="ds-status">No previewable slide dimensions are available. The source file is unchanged.</p>
 const active=result?.deck===deck&&result.index===index?result:undefined,geometry=active?.geometry
 const issues=geometry?[...geometry.originals.values()].flatMap(element=>{const issue=element.kind==='table'&&element.graphicFrameLayout==='source-anchored-v1'&&element.compatibility.status==='preserveOnly'?undefined:previewIssue(element,deck.assets);return issue?[`${element.name||element.kind}: ${issue}`]:[]}):[]
 if(geometry)for(const [id,message] of geometry.textIssues)issues.push(`${geometry.originals.get(id)?.name||id}: ${message}`)
 if(geometry?.omitted)issues.push(`Preview limited to 500 objects; ${geometry.omitted} remaining objects are preserved.`)
 return <section aria-label="Presentation file preview" className="ds-panel">
  <div className="ds-workstrip"><DsButton variant="outlined" disabled={index===0} onClick={()=>setAt(index-1)}>Previous slide</DsButton><DsField label="Slide"><DsSelect value={index} onChange={event=>setAt(Number(event.target.value))}>{deck.slides.map((item,i)=><option key={item.id} value={i}>{i+1} of {deck.slides.length}</option>)}</DsSelect></DsField><DsButton variant="outlined" disabled={index+1>=deck.slides.length} onClick={()=>setAt(index+1)}>Next slide</DsButton></div>
  <p className="ds-status">Approximate file preview · browser fonts and text wrapping. Not PowerPoint-equivalent rendering. Editing remains limited to verified targets below.</p>
  {geometry?<PptxFilePreviewVector deck={deck} geometry={geometry}/>:<p role="status">{active?.error?`Geometry preview unavailable: ${active.error}. The source file is unchanged.`:'Preparing local geometry preview…'}</p>}
  {issues.length>0&&<details open><summary>{issues.length} preview limitations</summary><ul>{issues.map((issue,i)=><li key={i}>{issue}</li>)}</ul></details>}
  {geometry?.tree.diagnostics.some(d=>d.code==='geometry.deterministicPathTone')&&<p role="note">Shaded paths use the native preview’s declared relative-tone policy; PowerPoint color equivalence is not established.</p>}
  {slide.compatibility.diagnostics.length>0&&<details><summary>Source compatibility warnings ({slide.compatibility.diagnostics.length})</summary><ul>{slide.compatibility.diagnostics.map((item,i)=><li key={i}>{item.message}</li>)}</ul></details>}
 </section>
}
