import type {PreviewNode,PreviewRect} from '../../../pptx-page-paint-worker/src/contract'

// A uniform change of SVG user units. Linear matrix terms and arc flags are
// dimensionless; every length, including glyph-local coordinates, is rescaled.
// Keep the validated worker transport in EMU; this is browser presentation only.
export const SVG_EMU_PER_POINT=12700
const length=(value:number)=>value/SVG_EMU_PER_POINT
const rect=(value:PreviewRect):PreviewRect=>({x:length(value.x),y:length(value.y),cx:length(value.cx),cy:length(value.cy)})
export function nativePptxSvgPath(d:string):string {
 const tokens=d.match(/[MLQCAZ]|[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/g)??[]
 const result:string[]=[]
 for(let i=0;i<tokens.length;){
  const command=tokens[i++]!,arity=({M:2,L:2,Q:4,C:6,A:7,Z:0} as Record<string,number>)[command]
  if(arity===undefined||i+arity>tokens.length)throw new TypeError('Invalid validated native SVG path.')
  result.push(command)
  for(let at=0;at<arity;at++){
   const value=Number(tokens[i++]!)
   if(!Number.isFinite(value))throw new TypeError('Invalid validated native SVG coordinate.')
   result.push(String(command==='A'&&at>=2&&at<=4?value:length(value)))
  }
 }
 return result.join(' ')
}
export function nativePptxSvgNode(node:PreviewNode):PreviewNode {
 switch(node.kind){
  case 'group':return {...node,transform:[...node.transform.slice(0,4),length(node.transform[4]),length(node.transform[5])] as [number,number,number,number,number,number],clip:node.clip?{...rect(node.clip),...(node.clip.radius===undefined?{}:{radius:length(node.clip.radius)}),...(node.clip.d===undefined?{}:{d:nativePptxSvgPath(node.clip.d)})}:undefined,children:node.children.map(nativePptxSvgNode)}
  case 'path':return {...node,d:nativePptxSvgPath(node.d),strokeWidth:node.strokeWidth===undefined?undefined:length(node.strokeWidth)}
  case 'rect':return {...node,rect:rect(node.rect),radius:length(node.radius),strokeWidth:node.strokeWidth===undefined?undefined:length(node.strokeWidth)}
  case 'ellipse':return {...node,rect:rect(node.rect),strokeWidth:node.strokeWidth===undefined?undefined:length(node.strokeWidth)}
  case 'image':case 'placeholder':return {...node,rect:rect(node.rect)}
 }
}
