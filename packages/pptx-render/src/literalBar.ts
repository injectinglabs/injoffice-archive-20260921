import type {NativeLiteralBar as LiteralBarRecord} from '@injoffice/pptx-native'
import {createCartesianBarPaths,type CartesianBarVector} from './cartesianBarPaths.js'
export type LiteralBarVector=CartesianBarVector
/** Literal source qualification stays separate from workbook-backed data. */
export function createNativeLiteralBarPaths(chart:LiteralBarRecord,cx:number,cy:number):readonly LiteralBarVector[]{
 if(!Number.isSafeInteger(cx)||!Number.isSafeInteger(cy)||cx<1||cy<1||cx>281474976710655||cy>281474976710655)throw new RangeError('invalid literal bar frame')
 if(chart.profile!=='literal-bar-v1'||chart.grouping!=='clustered'||chart.dataOrigin!=='literal'||!['column','bar'].includes(chart.barDirection)||!Number.isInteger(chart.gapWidth)||chart.gapWidth<0||chart.gapWidth>500||chart.overlap!==0||chart.categories.length<1||chart.categories.length>256||chart.categories.some(c=>typeof c!=='string')||chart.categories.reduce((sum,c)=>sum+c.length,0)>32768||chart.series.length<1||chart.series.length>16)throw new RangeError('invalid literal bar profile')
 return createCartesianBarPaths(chart,cx,cy)
}
