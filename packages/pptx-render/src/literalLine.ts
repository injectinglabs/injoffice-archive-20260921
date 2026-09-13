import type {NativeLiteralConnected} from '@injoffice/pptx-native'
import {createCartesianConnectedPaths,type CartesianConnectedVector} from './cartesianConnectedPaths.js'
export type LiteralConnectedVector=CartesianConnectedVector
export function createNativeConnectedLinePaths(chart:NativeLiteralConnected,cx:number,cy:number):readonly LiteralConnectedVector[]{
 if(!Number.isSafeInteger(cx)||!Number.isSafeInteger(cy)||cx<1||cy<1||cx>281474976710655||cy>281474976710655||chart.dataOrigin!=='literal'||!['literal-line-v1','literal-scatter-v1'].includes(chart.profile)||chart.series.length<1||chart.series.length>16)throw new RangeError('invalid connected chart frame/profile')
 return createCartesianConnectedPaths(chart,chart.profile==='literal-scatter-v1'?'scatter':'line',cx,cy)
}
export function createNativeLiteralLinePaths(chart:NativeLiteralConnected,cx:number,cy:number):readonly LiteralConnectedVector[]{
 if(chart.profile!=='literal-line-v1')throw new RangeError('expected literal line profile')
 return createNativeConnectedLinePaths(chart,cx,cy)
}
