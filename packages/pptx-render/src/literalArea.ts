import {createCartesianSignedAreaPaths} from './cartesianSignedAreaPaths.js'
import {validNativeLiteralArea,type NativeLiteralBarAxis,type NativeLiteralArea} from '@injoffice/pptx-native'
import type {RenderPathCommand,RenderStroke} from './types.js'
import {createCartesianAreaPaths} from './cartesianAreaPaths.js'
import {chartRational as rational,chartRationalDecimal as decimal,chartRationalSubtract as subtract,chartRationalDivide as divide,chartRationalCoordinate as coordinate} from './chartRational.js'

export type LiteralAreaInput = NativeLiteralArea
export interface LiteralAreaVector {
 readonly seriesIndex?:number;readonly axis?:'x'|'y'
 readonly path:readonly RenderPathCommand[];readonly color?:string;readonly stroke?:RenderStroke
}
const rgb=/^#[0-9A-F]{6}$/
function axis(a:NativeLiteralBarAxis):void{
 if(!a||!Number.isSafeInteger(a.id)||a.id<0||a.id>4294967295||!Number.isSafeInteger(a.crossAxisId)||a.crossAxisId<0||a.crossAxisId>4294967295||!['minMax','maxMin'].includes(a.orientation)||typeof a.deleted!=='boolean'||(a.deleted?(a.color!==undefined||a.widthEmu!==undefined||a.labels!==undefined):(typeof a.color!=='string'||rgb.exec(a.color)?.[0]!==a.color||!Number.isSafeInteger(a.widthEmu)||a.widthEmu!<1||a.widthEmu!>20116800)))throw new RangeError('invalid area axis style')
}

/** Literal data only. One compound fill per entire source series: artificial
 * interval edges never become separately antialiased/stroked shape boundaries.
 * Standard multi-series overlap follows XML series sequence as an explicit
 * preview policy, without claiming Office painter-order parity. Axis labels are
 * handled by the separately qualified shared text-layout integration. */
export function createNativeLiteralAreaPaths(chart:LiteralAreaInput,cx:number,cy:number):readonly LiteralAreaVector[]{
 if(!chart||chart.profile!=='literal-area-v1'||chart.dataOrigin!=='literal'||!Array.isArray(chart.categories)||chart.categories.length<1||chart.categories.length>256||!Array.isArray(chart.series)||chart.series.length<1||chart.series.length>16)throw new RangeError('invalid literal area profile')
 let units=0
 for(const category of chart.categories){if(typeof category!=='string')throw new RangeError('invalid area category');units+=category.length;if(units>32768)throw new RangeError('area category budget exceeded')}
 if(chart.sourceBaseline!==undefined&&!validNativeLiteralArea(chart))throw new RangeError('invalid source area baseline')
 const x=chart.xAxis,y=chart.yAxis
 axis(x);axis(y)
 if(x.id===y.id||x.crossAxisId!==y.id||y.crossAxisId!==x.id||x.position!=='b'||y.position!=='l'||x.min!==undefined||x.max!==undefined||x.crossesAt!==undefined||typeof y.min!=='string'||typeof y.max!=='string'||typeof y.crossesAt!=='string'||decimal(y.crossesAt).numerator!==0n)throw new RangeError('invalid area axis pairing')
 for(const s of chart.series)if(!s||!Array.isArray(s.values)||s.values.length!==chart.categories.length||typeof s.color!=='string'||rgb.exec(s.color)?.[0]!==s.color||(s.title!==undefined&&(typeof s.title!=='string'||s.title.length>1024)))throw new RangeError('invalid area series style/alignment')
 const signed=chart.grouping!=='standard'&&chart.series.some(s=>s.values.some((v:string)=>decimal(v).numerator<0n))
 if(signed&&!chart.sourceBaseline)throw new RangeError('signed area requires source minimum baseline')
 const data=signed?createCartesianSignedAreaPaths(chart.series,chart.grouping,{min:y.min,max:y.max,crossing:'min',reverseX:x.orientation==='maxMin',reverseY:y.orientation==='maxMin'},cx,cy):createCartesianAreaPaths(chart.series,chart.grouping,{min:y.min,max:y.max,...(chart.sourceBaseline?{baseline:chart.sourceBaseline.value}:{}),reverseX:x.orientation==='maxMin',reverseY:y.orientation==='maxMin'},cx,cy)
 const vectors:LiteralAreaVector[]=[]
 data.forEach((geometry,index)=>{
  const path=geometry.rings.flat()
  if(path.length>1536)throw new RangeError('generated area command budget exceeded')
  if(path.length) vectors.push({seriesIndex:geometry.seriesIndex,path,color:chart.series[index]!.color})
 })
 const minimum=decimal(y.min),span=subtract(decimal(y.max),minimum),baseline=coordinate(divide(subtract(chart.sourceBaseline?decimal(chart.sourceBaseline.value):rational(0n),minimum),span),cy,y.orientation==='minMax')
 if(!x.deleted)vectors.push({axis:'x',path:[{kind:'moveTo',x:0,y:baseline},{kind:'lineTo',x:cx,y:baseline}],stroke:{color:x.color!,widthEmu:x.widthEmu!,cap:'flat',dash:'solid'}})
 if(!y.deleted){const left=x.orientation==='maxMin'?cx:0;vectors.push({axis:'y',path:[{kind:'moveTo',x:left,y:0},{kind:'lineTo',x:left,y:cy}],stroke:{color:y.color!,widthEmu:y.widthEmu!,cap:'flat',dash:'solid'}})}
 return vectors
}
