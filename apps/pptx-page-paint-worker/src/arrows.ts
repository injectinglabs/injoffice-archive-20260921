import type {NativeArrowEnd} from '@injoffice/pptx-native'
import type {PreviewNode} from './contract.js'

export function previewArrowShaftInset(end:Readonly<NativeArrowEnd>|undefined,strokeWidth:number):number {
 if(!end||end.type==='none'||end.type==='arrow')return 0
 const length=strokeWidth*({sm:2,med:3,lg:5}[end.len??'med'])
 return length*(end.type==='triangle'?1:end.type==='stealth'?.75:.5)
}

/** Explicit InjOffice geometry, not PowerPoint-equivalent arrow metrics. */
export function previewArrow(end:Readonly<NativeArrowEnd>,tip:{x:number;y:number},direction:{x:number;y:number},strokeWidth:number,color:string):PreviewNode|undefined {
 if(end.type==='none')return undefined
 const magnitude=Math.hypot(direction.x,direction.y)
 if(!Number.isFinite(magnitude)||magnitude===0||!Number.isFinite(strokeWidth)||strokeWidth<=0)throw new Error('Arrow needs a nonzero bounded straight line and stroke')
 const factors={sm:2,med:3,lg:5},length=strokeWidth*factors[end.len??'med'],width=strokeWidth*factors[end.w??'med']
 if(Math.max(length,width)>1e9)throw new Error('Arrow geometry exceeds native preview budget')
 let node:PreviewNode
 switch(end.type){
  case 'oval':node={kind:'ellipse',rect:{x:-length,y:-width/2,cx:length,cy:width},fill:color};break
  case 'triangle':node={kind:'path',d:`M0 0 L${-length} ${-width/2} L${-length} ${width/2} Z`,fill:color};break
  case 'stealth':node={kind:'path',d:`M0 0 L${-length} ${-width/2} L${-length*.75} 0 L${-length} ${width/2} Z`,fill:color};break
  case 'diamond':node={kind:'path',d:`M0 0 L${-length/2} ${-width/2} L${-length} 0 L${-length/2} ${width/2} Z`,fill:color};break
  case 'arrow':node={kind:'path',d:`M${-length} ${-width/2} L0 0 L${-length} ${width/2}`,fill:'none',stroke:color,strokeWidth,strokeLinecap:'butt',strokeLinejoin:'miter',strokeMiterlimit:4};break
 }
 const x=direction.x/magnitude,y=direction.y/magnitude
 return {kind:'group',sourceRole:'connectorArrow',transform:[x,y,-y,x,tip.x,tip.y],children:[node]}
}
