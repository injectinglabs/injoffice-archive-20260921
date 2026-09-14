import {sourceTextBounds} from './sourceTextBounds.js'
import {SourceAffineBudget} from './sourceAffine.js'
import type {NativePptxTextLayout,RenderNode,RenderRect} from './types.js'

/** Complete generated chart surface envelope before clipping. Collect first,
 * then qualify one union in world coordinates so large bubble datasets share
 * the same affine budget. Child layout translations do not scale glyphs. */
export async function sourceGraphicFramePaintBounds(root:RenderNode,extents:NativePptxTextLayout['glyphExtents'],budget:SourceAffineBudget):Promise<RenderRect>{
 let left=Infinity,top=Infinity,right=-Infinity,bottom=-Infinity
 const point=(x:number,y:number,pad=0)=>{
  budget.charge(1)
  if([x,y,pad,x-pad,y-pad,x+pad,y+pad].some(v=>!Number.isSafeInteger(v))||pad<0)throw new RangeError('Invalid graphic-frame paint coordinate')
  left=Math.min(left,x-pad);top=Math.min(top,y-pad);right=Math.max(right,x+pad);bottom=Math.max(bottom,y+pad)
 }
 const rectangle=(r:RenderRect,x:number,y:number,pad=0)=>{point(x+r.x,y+r.y,pad);point(x+r.x+r.cx,y+r.y+r.cy,pad)}
 const visit=async(node:RenderNode,x:number,y:number,depth:number):Promise<void>=>{
  if(depth>64)throw new RangeError('Graphic-frame paint depth exceeded')
  budget.charge(8)
  if(node!==root){
   const t=node.transform
   if(t.sourceAffine||t.aPpm!==1000000||t.dPpm!==1000000||t.bPpm!==0||t.cPpm!==0||!Number.isSafeInteger(t.txEmu)||!Number.isSafeInteger(t.tyEmu))throw new RangeError('Graphic-frame layout children require exact translations')
   x+=t.txEmu;y+=t.tyEmu
  }
  if(node.clip)rectangle(node.clip.rect,x,y)
  if(node.kind==='group'){for(const child of node.children)await visit(child,x,y,depth+1);return}
  if(node.kind==='shape'||node.kind==='connector'){
   const stroke=node.stroke
   // A diagonal square cap reaches sqrt(2) half-widths in a coordinate
   // direction. Use the exact conservative factor 2 independently of joins.
   const joinFactor=stroke?.join==='round'||stroke?.join==='bevel'?1:stroke?.miterLimit??4
   const pad=stroke?Math.ceil(stroke.widthEmu/2*Math.max(joinFactor,stroke.cap==='square'?2:1)):0
   for(const command of node.path){
    if(command.kind==='close')continue
    if('rect' in command){rectangle(command.rect,x,y,pad);continue}
    point(x+command.x,y+command.y,pad)
    if('x1' in command)point(x+command.x1,y+command.y1,pad)
    if('x2' in command)point(x+command.x2,y+command.y2,pad)
    // Every point of an unrotated ellipse is within two radii of either
    // endpoint; this includes complete arcs and avoids implicit clipping.
    if(command.kind==='arcTo'){point(x+command.x-2*command.rx,y+command.y-2*command.ry,pad);point(x+command.x+2*command.rx,y+command.y+2*command.ry,pad)}
   }
   if(node.kind==='shape'&&node.textBody)throw new RangeError('Generated chart shape text needs separate layout qualification')
  }else if(node.kind==='text'){
   const body=node.textBody
   if(body.orientationTransform||body.transform||body.paragraphs.some(p=>p.transform))throw new RangeError('Chart labels require unscaled local glyph layout')
   rectangle(await sourceTextBounds({...body,bounds:{x:0,y:0,cx:0,cy:0}},extents),x,y)
  }else rectangle(node.bounds,x,y)
 }
 await visit(root,0,0,0)
 if(left===Infinity)return {x:0,y:0,cx:0,cy:0}
 if(!Number.isSafeInteger(right-left)||!Number.isSafeInteger(bottom-top))throw new RangeError('Graphic-frame paint hull exceeds precision')
 return {x:left,y:top,cx:right-left,cy:bottom-top}
}
