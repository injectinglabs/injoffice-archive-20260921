import {expect,it} from 'vitest'
import {validNativeLiteralRadar} from './chartRadarValidation.js'
export const radarFixture=()=>({profile:'literal-radar-v1',dataOrigin:'literal',style:'standard',categories:['A','B','C','D'],series:[{index:5,order:0,values:['2','4','6','8'],color:'#123456',widthEmu:12700}],categoryAxis:{id:10,crossAxisId:20,orientation:'minMax',position:'b',deleted:true},valueAxis:{id:20,crossAxisId:10,orientation:'minMax',position:'l',deleted:true,min:'0',max:'10',crossesAt:'0'}} as const)
it('qualifies dense source radial values, filled paint, signed scale and XML order permutation',()=>{
 expect(validNativeLiteralRadar(radarFixture())).toBe(true)
 const f=structuredClone(radarFixture()) as any;f.style='filled';f.series[0].fill='#ABCDEF';f.valueAxis.min='-10';f.series[0].values=['-5','0','5','10'];f.series.unshift({...f.series[0],index:9,order:1})
 expect(validNativeLiteralRadar(f)).toBe(true);expect(f.series.map((s:any)=>s.order)).toEqual([1,0])
})
it('refuses malformed source records, unqualified labels and outside-scale points',()=>{
 for(const mutate of [(f:any)=>f['']=0,(f:any)=>f.series[0]['']=0,(f:any)=>delete f.series[0],(f:any)=>delete f.series[0].values[0],(f:any)=>f.series[0].order=-0,(f:any)=>f.series[0].color='#123456\n',(f:any)=>f.series[0].values[0]='-1',(f:any)=>f.series[0].values[0]='11',(f:any)=>f.series[0].fill=undefined,(f:any)=>f.categoryAxis.labels={},(f:any)=>f.style='marker',(f:any)=>f.valueAxis.min='10']){
 const f=structuredClone(radarFixture());mutate(f);expect(validNativeLiteralRadar(f)).toBe(false)
 }
})
