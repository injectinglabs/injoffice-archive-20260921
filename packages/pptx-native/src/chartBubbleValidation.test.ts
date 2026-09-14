import {expect,it} from 'vitest'
import {validNativeLiteralBubble as valid} from './chartBubbleValidation.js'
const fixture=()=>({profile:'literal-bubble-v1',dataOrigin:'literal',bubbleScale:100,sizeRepresents:'area',series:[{index:7,order:0,xValues:['1','2'],values:['3','4'],sizes:['1e-100','4'],colors:['#123456','#ABCDEF']}],xAxis:{id:1,crossAxisId:2,orientation:'minMax',position:'b',deleted:true,min:'0',max:'10',crossesAt:'0'},yAxis:{id:2,crossAxisId:1,orientation:'minMax',position:'l',deleted:true,min:'0',max:'10',crossesAt:'0'}})
it('accepts explicit literal source and scale endpoints without source rewriting',()=>{
 const data=fixture(),before=JSON.stringify(data);expect(valid(data)).toBe(true);expect(JSON.stringify(data)).toBe(before)
 for(const bubbleScale of [0,300])expect(valid({...data,bubbleScale,sizeRepresents:'w'})).toBe(true)
})
it('rejects source origin, canonical integer, alignment, negative size and axis drift',()=>{
 const changes:((v:any)=>void)[]=[v=>v.dataOrigin='embedded-workbook',v=>v.series[0].valueReference={},v=>v.bubbleScale=-0,v=>v.bubbleScale=301,v=>v.series[0].order=-0,v=>v.series[0].sizes[0]='-1',v=>delete v.series[0].sizes[0],v=>v.series[0].xValues=['1'],v=>v.series[0].colors[0]='#abcdef',v=>v.yAxis.min='11',v=>v.xAxis.id=2,v=>v.yAxis.crossesAt='1',v=>v.yAxis.labels={},v=>v.series.push(v.series[0])]
 for(const change of changes){const data=fixture();change(data);expect(valid(data)).toBe(false)}
})

it('keeps closed records closed even for an empty own property name',()=>{
 const root:any=fixture();root['']=true;expect(valid(root)).toBe(false)
 const axis:any=fixture();axis.yAxis.deleted=false;axis.yAxis.color='#000000';axis.yAxis.widthEmu=12700;axis.yAxis.labels={profile:'explicit-axis-labels-v1',position:'low',majorTickMark:'none',majorUnit:'1',numberFormat:'0',style:{fontFamily:'DejaVu Sans',fontSize:1200,color:'123456',bold:false,italic:false,language:'en-US'}}
 expect(valid(axis)).toBe(true)
 for(const target of ['labels','style']){const copy=structuredClone(axis);(target==='labels'?copy.yAxis.labels:copy.yAxis.labels.style)['']=true;expect(valid(copy)).toBe(false)}
})
