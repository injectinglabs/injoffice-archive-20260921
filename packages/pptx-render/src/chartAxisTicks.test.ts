import {expect,it} from 'vitest'
import {createChartAxisTicks as create} from './chartAxisTicks.js'
it('generates signed explicit grids without binary decimal drift',()=>{
 const ticks=create('-1','.9','.5','0.0');expect(ticks.map(t=>t.label)).toEqual(['-1.0','-0.5','0.0','0.5']);expect(ticks[1]!.position).toEqual({numerator:5n,denominator:19n})
 expect(create('10000000000000000000000000000000','10000000000000000000000000000002','1','0').map(t=>t.label)).toEqual(['10000000000000000000000000000000','10000000000000000000000000000001','10000000000000000000000000000002'])
})
it('bounds tiny and excessive grids before iteration',()=>{
 expect(create('0','255e-100','1e-100','0.000000')).toHaveLength(256)
 for(const args of [['0','256','1'],['-.1','1','.3'],['0','1','0'],['0','1','-1'],['1','1','1'],['0','1e100','1e-100']])expect(()=>create(args[0]!,args[1]!,args[2]!,'0')).toThrow()
})
