import {expect,it} from 'vitest'
import {sourceGraphicFramePaintBounds} from './sourceGraphicFramePaintBounds.js'
import {SourceAffineBudget} from './sourceAffine.js'
import type {RenderShapeNode,NativePptxTextLayout} from './types.js'
const extents:NativePptxTextLayout['glyphExtents']=async()=>{throw new Error('No text expected')}
const shape:RenderShapeNode={kind:'shape',sourceElementId:'path',sourceKind:'chart',zIndex:0,compatibility:'preserveOnly',bounds:{x:0,y:0,cx:100,cy:100},transform:{aPpm:1000000,bPpm:0,cPpm:0,dPpm:1000000,txEmu:0,tyEmu:0},path:[{kind:'moveTo',x:0,y:0},{kind:'lineTo',x:100,y:100}],stroke:{color:'000000',widthEmu:20,cap:'square',join:'round'}}
it('encloses diagonal square caps independently of round or bevel joins',async()=>{
 for(const join of ['round','bevel'] as const){
  const hull=await sourceGraphicFramePaintBounds({...shape,stroke:{...shape.stroke!,join}},extents,new SourceAffineBudget())
  expect(hull).toEqual({x:-20,y:-20,cx:140,cy:140})
  // Exact diagonal square cap extends sqrt(2)*10 from its endpoint.
  expect(hull.x).toBeLessThan(-Math.SQRT2*10)
  expect(hull.x+hull.cx).toBeGreaterThan(100+Math.SQRT2*10)
 }
})
it('encloses curve control hulls and complete unrotated arc envelopes before clipping',async()=>{
 const path:RenderShapeNode={...shape,stroke:undefined,clip:{kind:'rect',rect:{x:0,y:0,cx:1,cy:1}},path:[{kind:'moveTo',x:0,y:0},{kind:'cubicBezierTo',x1:-200,y1:-300,x2:400,y2:500,x:100,y:100},{kind:'arcTo',rx:300,ry:100,largeArc:false,clockwise:true,x:100,y:100}]}
 expect(await sourceGraphicFramePaintBounds(path,extents,new SourceAffineBudget())).toEqual({x:-500,y:-300,cx:1200,cy:800})
})
it('refuses overflowing complete control envelopes instead of using the clipped frame',async()=>{
 await expect(sourceGraphicFramePaintBounds({...shape,path:[{kind:'moveTo',x:Number.MAX_SAFE_INTEGER,y:0}]},extents,new SourceAffineBudget())).rejects.toThrow('coordinate')
})
