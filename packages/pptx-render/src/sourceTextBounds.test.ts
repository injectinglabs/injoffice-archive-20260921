import {expect,it} from 'vitest'
import {sourceTextBounds} from './sourceTextBounds.js'
import type {RenderTextBodyNode} from './types.js'
it('includes the actual refused-run placeholder beyond its zero-advance paragraph hull',async()=>{
 const body={bounds:{x:0,y:0,cx:1,cy:1},paragraphs:[{x:0,y:0,widthEmu:0,heightEmu:0,runs:[{status:'refused',x:-100,fontSizeMilliPoints:12000,lineHeightEmu:200000}]}]} as unknown as RenderTextBodyNode
 expect(await sourceTextBounds(body,undefined)).toEqual({x:-100,y:0,cx:152400,cy:200000})
})
it('refuses a placeholder whose final endpoint is outside safe integer precision',async()=>{
 const body={bounds:{x:0,y:0,cx:1,cy:1},paragraphs:[{x:0,y:0,widthEmu:0,heightEmu:0,runs:[{status:'refused',x:Number.MAX_SAFE_INTEGER,fontSizeMilliPoints:12000,lineHeightEmu:200000}]}]} as unknown as RenderTextBodyNode
 await expect(sourceTextBounds(body,undefined)).rejects.toThrow('integer precision')
})
