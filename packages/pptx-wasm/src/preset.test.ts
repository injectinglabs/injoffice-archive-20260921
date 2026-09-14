import { expect, it } from 'vitest'
import { snapshotPresetRequest, decodePresetGeometry } from './preset.js'
const request={name:'triangle',widthEmu:1000,heightEmu:1000}
const geometry={profile:'drawingml-paths-v1',textRect:{x:250,y:500,cx:500,cy:500},paths:[{fillMode:'norm',stroke:true,commands:[{kind:'moveTo',x:500,y:0},{kind:'lineTo',x:1000,y:1000},{kind:'lineTo',x:0,y:1000},{kind:'close'}]}]}
it('decodes authored preserve-only geometry using the public native contract validator',()=>{
 const payload=snapshotPresetRequest(request)
 expect(decodePresetGeometry(JSON.stringify({protocol:'pptx-preset-evaluation-v1',request:{...request,adjustments:{}},geometry}),payload)).toEqual(geometry)
})
it('snapshots adjustments without invoking accessors or accepting unsafe values',()=>{
 const adjustments={adj:1000},input={...request,adjustments},payload=snapshotPresetRequest(input)
 adjustments.adj=2000;expect(JSON.parse(payload).adjustments.adj).toBe(1000)
 let calls=0;const accessor={...request,get adjustments(){calls++;return {}}}
 expect(()=>snapshotPresetRequest(accessor)).toThrow(/accessors/);expect(calls).toBe(0)
 for(const bad of [{...request,widthEmu:0},{...request,heightEmu:Infinity},{...request,adjustments:{adj:-0}},{...request,adjustments:null},{...request,unknown:1}])expect(()=>snapshotPresetRequest(bad as typeof request)).toThrow()
})
it('refuses request mismatch and malformed path responses',()=>{
 const payload=snapshotPresetRequest(request)
 const response={protocol:'pptx-preset-evaluation-v1',request:{...request,adjustments:{}},geometry}
 expect(()=>decodePresetGeometry(JSON.stringify({...response,request:{...response.request,widthEmu:2000}}),payload)).toThrow(/match/)
 expect(()=>decodePresetGeometry(JSON.stringify({...response,geometry:{...geometry,paths:[{fillMode:'norm',stroke:true,commands:[{kind:'arcTo',rx:1,ry:1,x:100,y:100}]}]}}),payload)).toThrow(/geometry/)
})
