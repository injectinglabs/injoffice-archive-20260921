import {expect,it} from 'vitest'
import {SourceAffineBudget,sourceHierarchyAffine} from './sourceAffine.js'
import {sourceTableTextBounds} from './sourceTableTextBounds.js'
import type {RenderTextBodyNode} from './types.js'
const world=()=>sourceHierarchyAffine({x:0,y:0,cx:100,cy:100},[],new SourceAffineBudget())
const body=()=>({kind:'textBody',status:'ok',bounds:{x:0,y:0,cx:100,cy:100},horizontalOverflow:'clip',paragraphs:[{x:0,y:0,widthEmu:0,heightEmu:0,transform:{aPpm:1000000,bPpm:0,cPpm:0,dPpm:1000000,txEmu:0,tyEmu:-20},runs:[{status:'refused',x:-10,fontSizeMilliPoints:10,lineHeightEmu:200}]}]}) as unknown as RenderTextBodyNode
it('keeps fixed cell x boundaries while including translated refused ink above and below',async()=>{
 const value=body(),snapshot=JSON.stringify(value),result=await sourceTableTextBounds(value,100,world(),undefined,1000,new SourceAffineBudget())
 expect(result.horizontalClip).toEqual({x:0,y:-21,cx:100,cy:202})
 expect(result.spans[0]!.bounds).toEqual({x:-10,y:0,cx:127,cy:200})
 expect(JSON.stringify(value)).toBe(snapshot)
})
it('matches body refusal paint, ignoring text-only transforms not replayed on its placeholder',async()=>{
 const value={...body(),status:'refused',orientationTransform:{aPpm:1000000,bPpm:0,cPpm:0,dPpm:1000000,txEmu:100000,tyEmu:100000}} as RenderTextBodyNode
 const result=await sourceTableTextBounds(value,100,world(),undefined,1000,new SourceAffineBudget())
 expect(result.horizontalClip).toEqual({x:0,y:-1,cx:100,cy:102})
})
it('checks actual transformed text ink even when horizontal overflow is allowed',async()=>{
 const value={...body(),horizontalOverflow:'overflow'} as RenderTextBodyNode
 await expect(sourceTableTextBounds(value,100,world(),undefined,100,new SourceAffineBudget())).rejects.toThrow()
})
it('uses supplied outline control hulls and rejects mismatched font identity',async()=>{
 const value=body(),paragraph=value.paragraphs[0]!
 const run={status:'ok',x:0,baselineY:0,fontSizeMilliPoints:10,faceId:'face',contentDigest:`sha256:${'a'.repeat(64)}`,glyphs:[{glyphId:7,xEmu:0,yEmu:0}]}
 const withGlyph={...value,paragraphs:[{...paragraph,runs:[run]}]} as unknown as RenderTextBodyNode
 const extents=()=>({faceId:'face',contentDigest:`sha256:${'a'.repeat(64)}` as const,glyphId:7,unitsPerEm:100,bounds:{xMin:-100,yMin:-100,xMax:100,yMax:200}})
 const result=await sourceTableTextBounds(withGlyph,100,world(),extents,1000,new SourceAffineBudget())
 expect(result.horizontalClip).toEqual({x:0,y:-275,cx:100,cy:383})
 await expect(sourceTableTextBounds(withGlyph,100,world(),()=>({...extents(),faceId:'wrong'}),1000,new SourceAffineBudget())).rejects.toThrow('identity')
})
