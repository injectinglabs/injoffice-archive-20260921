import {readFileSync} from 'node:fs'
import {expect,it} from 'vitest'
import type {NativePptxDeck} from '@injoffice/pptx-native'
import {createPptxWasmClient} from './index'
it('refuses custom geometry text mutation before starting a worker',()=>{
 const deck=JSON.parse(readFileSync(new URL('../../../go/pptxpatch/testdata/native-contract/valid/parsed-full.json',import.meta.url),'utf8')) as NativePptxDeck
 deck.sourceRevision=`rev-${'a'.repeat(64)}`
 const shape=deck.slides[0]!.elements.find(e=>e.kind==='shape')!
 if(shape.kind!=='shape'||!shape.source)throw Error('missing source shape')
 delete shape.preset
 shape.geometry={profile:'drawingml-paths-v1',textRect:{x:0,y:0,cx:100,cy:100},paths:[{fillMode:'norm',stroke:true,commands:[{kind:'moveTo',x:0,y:0},{kind:'lineTo',x:100,y:100}]}]}
 shape.compatibility={status:'preserveOnly',diagnostics:[{severity:'warning',code:'pptx.custom-geometry-preview',message:'Custom paths remain read-only'}]}
 let calls=0;const client=createPptxWasmClient({workerFactory:()=>{calls++;throw Error('worker must not start')}})
 expect(()=>client.apply(new Uint8Array([1]),deck,{expectedSourceRevision:deck.sourceRevision!,operations:[{operationId:'replace-custom-text',kind:'text.replace',elementId:shape.id,expectedFingerprintSha256:shape.source!.fingerprintSha256,paragraphs:[]}]})).toThrow('Custom geometry is preview-only')
 expect(calls).toBe(0)
})
