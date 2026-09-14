import {validateNativePptx,type NativeElement,type NativePptxDeck} from '@injoffice/pptx-native'
import {compileNativePptxSlide,createRecordingPaintSurface,paintSlideRenderTree,renderTransformMatrix,SourceAffineBudget,type NativePptxTextLayout,type PaintCommand,type RenderNode,type RenderPathCommand,type RenderTransform,type SlideRenderTree} from '@injoffice/pptx-render'

export const FILE_PREVIEW_EMU_PER_POINT=12700
export const filePreviewLength=(value:number)=>value/FILE_PREVIEW_EMU_PER_POINT
export const filePreviewMatrix=(value:RenderTransform,budget:SourceAffineBudget)=>{
 const [a,b,c,d,e,f]=renderTransformMatrix(value,budget)
 return [a,b,c,d,filePreviewLength(e),filePreviewLength(f)] as const
}
export function filePreviewPath(commands:readonly RenderPathCommand[]):string {
 const n=filePreviewLength
 return commands.map(c=>{
  switch(c.kind){
   case 'moveTo':return `M ${n(c.x)} ${n(c.y)}`
   case 'lineTo':return `L ${n(c.x)} ${n(c.y)}`
   case 'quadBezierTo':return `Q ${n(c.x1)} ${n(c.y1)} ${n(c.x)} ${n(c.y)}`
   case 'cubicBezierTo':return `C ${n(c.x1)} ${n(c.y1)} ${n(c.x2)} ${n(c.y2)} ${n(c.x)} ${n(c.y)}`
   case 'arcTo':return `A ${n(c.rx)} ${n(c.ry)} 0 ${c.largeArc?1:0} ${c.clockwise?1:0} ${n(c.x)} ${n(c.y)}`
   case 'close':return 'Z'
   default:throw new TypeError('Native path primitive requires an SVG primitive.')
  }
 }).join(' ')
}
// The local file preview has no operator-supplied font bytes. The private
// geometry projection contains zero runs, and must never resolve/shapes fonts.
const geometryOnlyLayout:NativePptxTextLayout={
 manifest:{version:1,manifestId:'local-geometry-only',revision:'1',faces:[{faceId:'unavailable',family:'Unavailable geometry-only placeholder',weight:400,style:'normal',stretch:100,source:{kind:'host',resourceId:'unavailable'}}],fallbackChains:[]},
 resolver:{providerId:'local-geometry-only',providerRevision:'1',resolve(){throw new Error('Geometry-only projection attempted font resolution')},load(){throw new Error('Geometry-only projection attempted font loading')}},
 shaper:{providerId:'local-geometry-only',providerRevision:'1',shape(){throw new Error('Geometry-only projection attempted text shaping')}},
 defaults:{fontFamilies:['sans-serif'],fontSizeHundredthPt:1800,script:'Latn',language:'en-US',direction:'ltr'},
}
export interface FilePreviewGeometry {
 readonly tree:SlideRenderTree
 readonly originals:ReadonlyMap<string,NativeElement>
 readonly paths:ReadonlyMap<string,readonly Extract<PaintCommand,{kind:'path'}>[]>
 readonly matrices:ReadonlyMap<RenderTransform,readonly number[]>
 readonly omitted:number
 readonly textIssues:ReadonlyMap<string,string>
}
/** Private render projection only: never return the deck as a source/mutation
 * model. Geometry, source authority, text-body metadata and resources survive;
 * original paragraphs stay in originals for explicitly approximate DOM text. */
export async function compileFilePreviewGeometry(deck:NativePptxDeck,index:number):Promise<FilePreviewGeometry>{
 const validation=validateNativePptx(deck)
 if(!validation.ok)throw new TypeError('File preview requires a validated native presentation.')
 if(!Number.isSafeInteger(index)||index<0||index>=deck.slides.length)throw new RangeError('File preview slide is unavailable.')
 const projection=structuredClone(deck),originals=new Map<string,NativeElement>()
 let count=0,omitted=0
 const textIssues=new Map<string,string>()
 const countTree=(element:NativeElement):number=>1+(element.kind==='group'?element.children.reduce((sum,e)=>sum+countTree(e),0):0)
 const select=(elements:NativeElement[],source:NativeElement[]):NativeElement[]=>elements.flatMap((element,i)=>{
  if(count>=500){omitted+=countTree(element);return []}
  count++;originals.set(element.id,source[i]!)
  if(element.kind==='group')element.children=select(element.children,(source[i] as Extract<NativeElement,{kind:'group'}>).children)
  if(element.kind==='shape'||element.kind==='text'){
   // Empty runs cannot establish the native vertical-script/list admission.
   // Retain that original-content boundary before the geometry projection.
   if(element.textBody?.writingMode==='vertical-clockwise'&&element.paragraphs.some(p=>p.bullet!==false||p.level!==0||(p.marginLeftEmu??0)!==0||(p.indentEmu??0)!==0||p.runs.some(r=>!r.text||!/^[\x20-\x7e]+$/.test(r.text))))textIssues.set(element.id,'Vertical text requires supported Latin content and paragraph semantics.')
   element.paragraphs=[]
  }
  if(element.kind==='table')for(let rowIndex=0;rowIndex<element.table.rows.length;rowIndex++)for(let columnIndex=0;columnIndex<element.table.rows[rowIndex]!.length;columnIndex++){
   const cell=element.table.rows[rowIndex]![columnIndex]!
   if(cell.textBody?.writingMode==='vertical-clockwise'&&cell.paragraphs?.some(p=>p.bullet!==false||p.level!==0||(p.marginLeftEmu??0)!==0||(p.indentEmu??0)!==0||p.runs.some(r=>!r.text||!/^[\x20-\x7e]+$/.test(r.text))))textIssues.set(`${element.id}-cell-${rowIndex*element.table.columnWidths.length+columnIndex}`,'Vertical table text requires supported Latin content and paragraph semantics.')
   cell.text=''
   if(cell.paragraphs)cell.paragraphs=[]
  }
  return [element]
 })
 projection.slides[index]!.elements=select(projection.slides[index]!.elements,deck.slides[index]!.elements)
 const tree=await compileNativePptxSlide(projection,index,{textLayout:geometryOnlyLayout})
 const recording=createRecordingPaintSurface();paintSlideRenderTree(tree,recording)
 const paths=new Map<string,Extract<PaintCommand,{kind:'path'}>[]>(),matrices=new Map<RenderTransform,readonly number[]>(),budget=new SourceAffineBudget()
 for(const command of recording.finish())if(command.kind==='path'){
  const list=paths.get(command.sourceElementId)??[];list.push(command);paths.set(command.sourceElementId,list)
 }
 const visit=(node:RenderNode)=>{
  matrices.set(node.transform,filePreviewMatrix(node.transform,budget))
  if(node.kind==='shape'||node.kind==='text'){
   if(node.textBody?.orientationTransform)matrices.set(node.textBody.orientationTransform,filePreviewMatrix(node.textBody.orientationTransform,budget))
   if(node.textBody?.transform)matrices.set(node.textBody.transform,filePreviewMatrix(node.textBody.transform,budget))
  }
  if(node.kind==='table')for(const cell of node.cells){
   if(cell.textBody?.orientationTransform)matrices.set(cell.textBody.orientationTransform,filePreviewMatrix(cell.textBody.orientationTransform,budget))
   if(cell.textBody?.transform)matrices.set(cell.textBody.transform,filePreviewMatrix(cell.textBody.transform,budget))
  }
  if(node.kind==='group')node.children.forEach(visit)
 }
 tree.nodes.forEach(visit)
 return {tree,originals,paths,matrices,omitted,textIssues}
}
