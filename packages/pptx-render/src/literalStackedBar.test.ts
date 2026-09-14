import {it,expect} from 'vitest'
import {createNativeLiteralStackedBarPaths,type LiteralStackedBarInput} from './literalStackedBar.js'
it('requires literal profile authority and retains raw source order',()=>{
 const chart:LiteralStackedBarInput={profile:'literal-stacked-bar-v1',dataOrigin:'literal',grouping:'percentStacked',overlap:100,gapWidth:0,barDirection:'column',categories:['A'],series:[{index:9,order:1,values:['-2'],colors:['#123456']},{index:7,order:0,values:['1'],colors:['#ABCDEF']}],categoryAxis:{id:1,crossAxisId:2,position:'b',orientation:'minMax',deleted:true},valueAxis:{id:2,crossAxisId:1,position:'l',orientation:'minMax',deleted:true,min:'-1',max:'1',crossesAt:'0'}}
 const before=JSON.stringify(chart);expect(createNativeLiteralStackedBarPaths(chart,600,600).map(p=>p.seriesIndex)).toEqual([9,7]);expect(JSON.stringify(chart)).toBe(before)
 for(const bad of [{...chart,profile:'literal-bar-v1'},{...chart,dataOrigin:'reference'}])expect(()=>createNativeLiteralStackedBarPaths(bad as LiteralStackedBarInput,600,600)).toThrow()
})
