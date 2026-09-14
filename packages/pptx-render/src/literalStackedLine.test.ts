import {it,expect} from 'vitest'
import {createNativeLiteralStackedLinePaths,type LiteralStackedLineInput} from './literalStackedLine.js'
it('requires literal line authority without normalizing source metadata',()=>{
 const chart:LiteralStackedLineInput={profile:'literal-stacked-line-v1',dataOrigin:'literal',grouping:'stacked',categories:['A','B'],series:[{index:9,order:1,values:['-2','0'],color:'#123456',widthEmu:10},{index:7,order:0,values:['1','0'],color:'#ABCDEF',widthEmu:10}],xAxis:{id:1,crossAxisId:2,position:'b',orientation:'minMax',deleted:true},yAxis:{id:2,crossAxisId:1,position:'l',orientation:'minMax',deleted:true,min:'-3',max:'3',crossesAt:'0'}}
 const before=JSON.stringify(chart);expect(createNativeLiteralStackedLinePaths(chart,600,600).map(p=>p.seriesIndex)).toEqual([9,7]);expect(JSON.stringify(chart)).toBe(before)
 for(const bad of [{...chart,profile:'literal-line-v1'},{...chart,dataOrigin:'reference'}])expect(()=>createNativeLiteralStackedLinePaths(bad as LiteralStackedLineInput,600,600)).toThrow()
})
