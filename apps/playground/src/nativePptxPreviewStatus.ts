import type {NativeDiagnostic,NativeElement,NativePptxDeck} from '@injoffice/pptx-native'
import type {PptxPreview,PreviewNode} from '../../pptx-page-paint-worker/src/contract'

export interface NativePptxPreviewStatus {
 status:'partial'|'unverified'|'unavailable'
 label:string
 paintPrimitives:number
 glyphRecords:number
 placeholderRegions:number
 knownRefusedObjects:number
 sourceBound:boolean
 reasons:string[]
}

// Paint records are not source objects. The v1 helper does not carry sufficient
// source coverage evidence to certify a complete slide, even with no diagnostics.
export function nativePptxPreviewStatus(preview:PptxPreview,source?:NativePptxDeck):NativePptxPreviewStatus {
 let paintPrimitives=0,glyphRecords=0,placeholderRegions=0
 const reasons=new Set<string>()
 const visibleStroke=(node:{stroke?:string;strokeWidth?:number})=>!!node.stroke&&node.stroke!=='none'&&(node.strokeWidth??0)>0
 function visit(node:PreviewNode){
  if(node.kind==='group'){
   if(node.sourceRole==='contentRun')glyphRecords++
   node.children.forEach(visit)
  }else if(node.kind==='placeholder'){
   placeholderRegions++;reasons.add(`Placeholder region: ${node.label}`)
  }else if(node.kind==='path'){
   if(node.d.trim()&&(node.fill!=='none'||visibleStroke(node)))paintPrimitives++
  }else if(node.kind==='image'){
   if(node.rect.cx>0&&node.rect.cy>0)paintPrimitives++
  }else if(node.rect.cx>0&&node.rect.cy>0&&(node.fill!=='none'||visibleStroke(node)))paintPrimitives++
 }
 preview.nodes.forEach(visit)
 const slide=source?.slides[preview.slide_index]
 const sourceBound=source?.origin==='parsed'&&source.sourceRevision===`rev-${preview.package_sha256}`&&source.slides.length===preview.slide_count&&!!slide
 const refused=new Set<string>()
 let sourceGap=false
 if(sourceBound&&source&&slide){
  const names=new Map<string,string>()
  const elements=(items:NativeElement[])=>{for(const element of items){
   names.set(element.id,element.name??element.id)
   if(element.compatibility.status==='refused')refused.add(element.id)
   if(element.kind==='group')elements(element.children)
  }}
  elements(slide.elements)
  const diagnostic=(item:NativeDiagnostic)=>{
   if(item.scope?.slideId&&item.scope.slideId!==slide.id)return
   if(item.scope?.elementId&&!names.has(item.scope.elementId)&&!item.scope.slideId)return
   if(item.severity==='refusal'||/unsupported|unavailable|refused/.test(item.code)){
    sourceGap=true
    if(item.scope?.elementId&&item.severity==='refusal')refused.add(item.scope.elementId)
    const target=item.scope?.elementId?(names.get(item.scope.elementId)??item.scope.elementId):`Slide ${preview.slide_index+1}`
    reasons.add(`${target}: ${item.message}`)
   }
  }
  source.compatibility.diagnostics.forEach(diagnostic)
  slide.compatibility.diagnostics.forEach(diagnostic)
  const elementDiagnostics=(items:NativeElement[])=>{for(const element of items){
   element.compatibility.diagnostics.forEach(item=>diagnostic({...item,scope:{...item.scope,slideId:slide.id,elementId:element.id}}))
   if(element.compatibility.status==='refused'&&!element.compatibility.diagnostics.length)reasons.add(`${names.get(element.id)}: source object is refused by the native extractor`)
   if(element.kind==='group')elementDiagnostics(element.children)
  }}
  elementDiagnostics(slide.elements)
 }else reasons.add('Source-object coverage could not be joined to this response; no source object counts are claimed.')
 let runtimeGap=false
 for(const message of preview.diagnostics){
  const code=message.slice(0,message.indexOf(':')<0?message.length:message.indexOf(':'))
  // Compatibility messages are provided with scope/severity by the joined deck.
  // Read-only preservation and deterministic-layout notices are not omissions.
  if(code==='native.compatibility'||code==='render.preserveOnly'||code==='text.deterministicLayout'||code==='arrow.deterministicGeometry')continue
  if(/unavailable|unsupported|refused|missing/i.test(code))runtimeGap=true
  reasons.add(message)
 }
 const status=paintPrimitives===0?'unavailable':placeholderRegions>0||refused.size>0||sourceGap||runtimeGap?'partial':'unverified'
 if(status==='unavailable')reasons.add('The response contains no non-placeholder painted content. A background or placeholder is not a rendered document.')
 return {status,label:status==='partial'?'Partial native preview':status==='unavailable'?'Native preview unavailable':'Native preview — coverage unverified',paintPrimitives,glyphRecords,placeholderRegions,knownRefusedObjects:refused.size,sourceBound,reasons:[...reasons]}
}
