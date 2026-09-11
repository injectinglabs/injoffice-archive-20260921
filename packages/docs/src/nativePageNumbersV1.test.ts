import { expect,it } from 'vitest'
import { nativeDocxPageNumberV1 } from './nativePageNumbersV1.js'
import type { NativeDocxDocumentV1 } from './nativeContract.js'
import type { NativeDocxPaginatedLayoutV1 } from './nativePaginationV1.js'

function sample(starts: Array<number|undefined>, ids: string[][]) {
 const document={sections:starts.map((start,index)=>({id:String(index),...(start===undefined?{}:{page_number_start:start})}))} as NativeDocxDocumentV1
 const layout={status:'paginated',pages:ids.map((section_ids,ordinal)=>({ordinal,kind:'content',section_ids}))} as NativeDocxPaginatedLayoutV1
 return {document,layout}
}
it('continues physical numbering and applies later decimal restarts including zero',()=>{
 const {document,layout}=sample([7,undefined,0],[['0'],['0'],['1'],['2'],['2']])
 expect(layout.status==='paginated'&&layout.pages.map((_,ordinal)=>nativeDocxPageNumberV1(document,layout,ordinal))).toEqual([7,8,9,0,1])
})
it('refuses shared-page continuous restarts and number overflow',()=>{
 const a=sample([1,7],[['0','1']]);expect(()=>nativeDocxPageNumberV1(a.document,a.layout,0)).toThrow(/shared/)
 const b=sample([999999],[['0'],['0']]);expect(()=>nativeDocxPageNumberV1(b.document,b.layout,1)).toThrow(/bounded/)
})
