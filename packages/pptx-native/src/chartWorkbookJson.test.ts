import {expect,it} from 'vitest'
import {parseChartWorkbookJson as parse} from './chartWorkbookJson.js'
it('preserves exact integer/string records and refuses lossy duplicate or numeric JSON',()=>{
 expect(parse('{"a":[true,false,null,12,"𐀀"],"b":{}}')).toEqual({a:[true,false,null,12,'𐀀'],b:{}})
 for(const text of ['{"a":1,"a":2}','{"a":1,"\\u0061":2}','{"a":{"x":1,"x":2}}','9007199254740993','-0','1.0','1e2','[1,]','{"a":1,}','"\\ud800"','true false'])expect(()=>parse(text)).toThrow()
})
