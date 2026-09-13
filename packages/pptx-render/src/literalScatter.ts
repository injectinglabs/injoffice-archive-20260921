import type {NativeLiteralConnected} from '@injoffice/pptx-native'
import {createNativeConnectedLinePaths,type LiteralConnectedVector} from './literalLine.js'

/** Connected XY scatter follows explicit source order, never sorted x order. */
export function createNativeLiteralScatterPaths(chart:NativeLiteralConnected,cx:number,cy:number):readonly LiteralConnectedVector[]{
 if(chart.profile!=='literal-scatter-v1')throw new RangeError('expected literal scatter profile')
 return createNativeConnectedLinePaths(chart,cx,cy)
}
