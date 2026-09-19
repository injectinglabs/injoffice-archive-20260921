import {expect,it} from 'vitest'
import {compoundStrokeBands,offsetRectanglePath} from './compoundStroke.js'
import {createRecordingPaintSurface,paintSlideRenderTree} from './paint.js'
import {PPTX_RENDER_TREE_VERSION} from './types.js'
import type {RenderPathCommand,RenderShapeNode,RenderStroke,SlideRenderTree} from './types.js'

const stroke=(compound:RenderStroke['compound'],widthEmu=63500):RenderStroke=>({color:'7030A0',widthEmu,cap:'flat',join:'miter',dash:'solid',compound,miterLimit:800000})
const corners=(x:number,y:number,cx:number,cy:number):readonly RenderPathCommand[]=>[
 {kind:'moveTo',x,y},{kind:'lineTo',x:x+cx,y},{kind:'lineTo',x:x+cx,y:y+cy},{kind:'lineTo',x,y:y+cy},{kind:'close'},
]
// What presetPath('rect') actually compiles to, and what a compound outline on
// a rect placeholder therefore has to offset.
const rect=(x:number,y:number,cx:number,cy:number):readonly RenderPathCommand[]=>[{kind:'rect',rect:{x,y,cx,cy}}]

it('paints a single line as the one centered stroke every caller already emits',()=>{
 expect(compoundStrokeBands(stroke(undefined))).toBeUndefined()
 expect(compoundStrokeBands(stroke('single'))).toBeUndefined()
})

// hard-v2 ShapeLineProperties.pptx: a thinThick outline of w=63500 EMU. Against
// PowerPoint's own raster of that slide the outline color runs 217.33..219.00 px
// and 220.67..224.00 px with a gap between, on a shape whose left edge is at
// 220.67 px — bands of 1:1:2 over four parts, centered on the outline.
it('splits thinThick the way PowerPoint rasterizes it',()=>{
 expect(compoundStrokeBands(stroke('thinThick'))).toEqual([
  {offsetEmu:-23812,widthEmu:15875},
  {offsetEmu:15875,widthEmu:31750},
 ])
})

it('orders every compound outward edge to inward edge across the authored width',()=>{
 for(const compound of ['double','thickThin','thinThick','triple'] as const){
  const bands=compoundStrokeBands(stroke(compound))!
  expect(bands.length).toBeGreaterThan(1)
  let previous=-Infinity
  for(const band of bands){
   expect(band.widthEmu).toBeGreaterThan(0)
   // Every band sits inside the authored width and never overlaps its
   // neighbour, up to the half EMU an odd band width cannot center on.
   expect(band.offsetEmu-band.widthEmu/2).toBeGreaterThanOrEqual(-63500/2-0.5)
   expect(band.offsetEmu+band.widthEmu/2).toBeLessThanOrEqual(63500/2+0.5)
   expect(band.offsetEmu-band.widthEmu/2).toBeGreaterThanOrEqual(previous-1)
   previous=band.offsetEmu+band.widthEmu/2
  }
 }
 // thickThin is thinThick mirrored: the thick line leads on the outside.
 const [outer,inner]=compoundStrokeBands(stroke('thickThin'))!
 expect(outer!.widthEmu).toBe(31750)
 expect(inner!.widthEmu).toBe(15875)
 expect(compoundStrokeBands(stroke('double'))!.map(b=>b.widthEmu)).toEqual([21167,21167])
 expect(compoundStrokeBands(stroke('triple'))!.map(b=>b.widthEmu)).toEqual([12700,12700,12700])
})

it('never emits a band that would paint nothing, however thin the outline',()=>{
 for(const compound of ['double','thickThin','thinThick','triple'] as const){
  for(const widthEmu of [0,1,2,3,5,7,11,12700]){
   for(const band of compoundStrokeBands(stroke(compound,widthEmu))??[]){
    expect(band.widthEmu).toBeGreaterThan(0)
    expect(Number.isSafeInteger(band.offsetEmu)).toBe(true)
   }
  }
 }
 // An outline too thin to separate collapses to no bands at all rather than to
 // a stack of zero-width strokes.
 expect(compoundStrokeBands(stroke('triple',0))).toBeUndefined()
})

it('moves every side of a rectangle the same distance, keeping winding and start corner',()=>{
 expect(offsetRectanglePath(rect(100,200,1000,600),50)).toEqual(rect(150,250,900,500))
 expect(offsetRectanglePath(rect(100,200,1000,600),-50)).toEqual(rect(50,150,1100,700))
 expect(offsetRectanglePath(corners(100,200,1000,600),50)).toEqual(corners(150,250,900,500))
 expect(offsetRectanglePath(corners(100,200,1000,600),-50)).toEqual(corners(50,150,1100,700))
 // Winding and starting corner survive: a rectangle wound the other way keeps it.
 expect(offsetRectanglePath([{kind:'moveTo',x:1100,y:800},{kind:'lineTo',x:100,y:800},{kind:'lineTo',x:100,y:200},{kind:'lineTo',x:1100,y:200},{kind:'close'}],50))
  .toEqual([{kind:'moveTo',x:1050,y:750},{kind:'lineTo',x:150,y:750},{kind:'lineTo',x:150,y:250},{kind:'lineTo',x:1050,y:250},{kind:'close'}])
})

it('refuses to offset anything that is not an axis-aligned rectangle',()=>{
 expect(offsetRectanglePath([{kind:'moveTo',x:0,y:0},{kind:'lineTo',x:100,y:10},{kind:'lineTo',x:100,y:100},{kind:'lineTo',x:0,y:100},{kind:'close'}],5)).toBeUndefined()
 expect(offsetRectanglePath([{kind:'moveTo',x:0,y:0},{kind:'lineTo',x:100,y:0},{kind:'lineTo',x:50,y:100},{kind:'close'}],5)).toBeUndefined()
 expect(offsetRectanglePath([{kind:'moveTo',x:0,y:0},{kind:'quadBezierTo',x1:50,y1:50,x:100,y:0},{kind:'lineTo',x:100,y:100},{kind:'lineTo',x:0,y:100},{kind:'close'}],5)).toBeUndefined()
 // An inset that would collapse the rectangle has no bands to paint.
 expect(offsetRectanglePath(rect(0,0,100,80),40)).toBeUndefined()
})

it('paints a compound outline as its bands and nothing else',()=>{
 const shape:RenderShapeNode={
  kind:'shape',sourceElementId:'element-1',sourceKind:'shape',zIndex:0,compatibility:'editable',
  transform:{aPpm:1_000_000,bPpm:0,cPpm:0,dPpm:1_000_000,txEmu:0,tyEmu:0},
  bounds:{x:0,y:0,cx:4857750,cy:1285875},path:rect(0,0,4857750,1285875),
  fill:{color:'FF0000'},stroke:stroke('thinThick'),
 }
 const tree:SlideRenderTree={
  version:PPTX_RENDER_TREE_VERSION,documentId:'d',slideId:'s',slideIndex:0,
  size:{cx:9144000,cy:6858000},background:{color:'FFFFFF'},clip:{kind:'rect',rect:{x:0,y:0,cx:9144000,cy:6858000}},
  nodes:[shape],assets:[],diagnostics:[],
 }
 const surface=createRecordingPaintSurface()
 paintSlideRenderTree(tree,surface)
 const paths=surface.finish().filter(command=>command.kind==='path')
 // The fill is painted once, unstroked, then one ordinary stroke per band. No
 // stroke of the full 63500 EMU width survives, and no band restates the
 // compound, so the paint stream and the preview transport stay unchanged.
 expect(paths.map(command=>({fill:command.fill,width:command.stroke?.widthEmu,path:command.path}))).toEqual([
  {fill:'FF0000',width:undefined,path:rect(0,0,4857750,1285875)},
  {fill:undefined,width:15875,path:rect(-23812,-23812,4857750+2*23812,1285875+2*23812)},
  {fill:undefined,width:31750,path:rect(15875,15875,4857750-2*15875,1285875-2*15875)},
 ])
 expect(paths.every(command=>command.stroke?.compound===undefined)).toBe(true)
})

it('paints a single outline exactly as it did before compound outlines existed',()=>{
 const base={kind:'shape',sourceElementId:'element-1',sourceKind:'shape',zIndex:0,compatibility:'editable',
  transform:{aPpm:1_000_000,bPpm:0,cPpm:0,dPpm:1_000_000,txEmu:0,tyEmu:0},
  bounds:{x:0,y:0,cx:1000,cy:800},path:rect(0,0,1000,800),fill:{color:'FF0000'}} as const
 const paint=(shapeStroke:RenderStroke)=>{
  const surface=createRecordingPaintSurface()
  paintSlideRenderTree({version:PPTX_RENDER_TREE_VERSION,documentId:'d',slideId:'s',slideIndex:0,size:{cx:9144000,cy:6858000},
   background:{color:'FFFFFF'},clip:{kind:'rect',rect:{x:0,y:0,cx:9144000,cy:6858000}},
   nodes:[{...base,stroke:shapeStroke}],assets:[],diagnostics:[]},surface)
  return surface.finish().filter(command=>command.kind==='path')
 }
 for(const compound of [undefined,'single'] as const){
  const paths=paint(stroke(compound,12700))
  expect(paths).toHaveLength(1)
  expect(paths[0]!.fill).toBe('FF0000')
  expect(paths[0]!.stroke?.widthEmu).toBe(12700)
 }
})
