import {expect,it} from 'vitest'
import {measureChartAxisText as measure} from './chartAxisText.js'
import type {RenderTextBodyNode,NativePptxGlyphExtents} from './types.js'
const body=()=>({kind:'textBody',sourceElementId:'label',bounds:{x:0,y:0,cx:10000,cy:10000},fidelity:'native',status:'laidOut',paragraphs:[{x:0,y:0,widthEmu:100,heightEmu:200,runs:[{status:'shaped',faceId:'face',contentDigest:'digest',fontSelection:{resolution:'exact'},fontSizeMilliPoints:1000,x:0,baselineY:100,glyphs:[{glyphId:7,xEmu:10,yEmu:-20}]}]}]} as unknown as RenderTextBodyNode)
const extent:NativePptxGlyphExtents={faceId:'face',contentDigest:'digest',glyphId:7,unitsPerEm:1000,bounds:{xMin:-2,yMin:-5,xMax:9,yMax:12}}
it('unions positioned ink hulls with advances and rounds outward',async()=>{
 // Scale=12.7EMU/design unit; font-up y reverses around baseline+offset=80.
 expect(await measure(body(),()=>extent)).toEqual({x:-16,y:-73,cx:141,cy:273})
 expect(await measure(body(),()=>({...extent,unitsPerEm:2000,bounds:{xMin:-4,yMin:-10,xMax:18,yMax:24}}))).toEqual({x:-16,y:-73,cx:141,cy:273})
})
it('requires exact font identity and bounded conservative extents',async()=>{
 await expect(measure(body(),undefined)).rejects.toThrow()
 for(const changed of [{faceId:'other'},{glyphId:8},{contentDigest:'other'},{unitsPerEm:0},{bounds:{xMin:1,yMin:0,xMax:0,yMax:1}},{bounds:{xMin:0,yMin:0,xMax:Infinity,yMax:1}}])await expect(measure(body(),()=>({...extent,...changed}))).rejects.toThrow()
 const substitute=body();(substitute.paragraphs[0]!.runs[0]!.fontSelection as {resolution:string}).resolution='substitute';await expect(measure(substitute,()=>extent)).rejects.toThrow()
 expect(await measure(body(),()=>({...extent,bounds:null}))).toEqual({x:0,y:0,cx:100,cy:200})
})
