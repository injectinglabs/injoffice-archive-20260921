import {createNativeConnectedLinePaths,type LiteralConnectedRecord,type LiteralConnectedVector} from './literalLine.js'

/** Connected XY scatter follows explicit source order, never sorted x order. */
export function createNativeLiteralScatterPaths(chart:LiteralConnectedRecord,cx:number,cy:number):readonly LiteralConnectedVector[]{
 if(chart.profile!=='literal-scatter-v1')throw new RangeError('expected literal scatter profile')
 return createNativeConnectedLinePaths(chart,cx,cy)
}
