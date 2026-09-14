import {createCartesianStackedLinePaths,type CartesianStackedLineInput} from './cartesianStackedLinePaths.js'
export type {CartesianConnectedVector as LiteralStackedLineVector} from './cartesianConnectedPaths.js'
export type LiteralStackedLineInput=CartesianStackedLineInput&{readonly profile:'literal-stacked-line-v1';readonly dataOrigin:'literal'}
/** Literal authority is never inferred from a reference/cache record. The caller
 * keeps original XML sequence and order metadata; no source object is rewritten. */
export function createNativeLiteralStackedLinePaths(chart:LiteralStackedLineInput,cx:number,cy:number){
 if(chart.profile!=='literal-stacked-line-v1'||chart.dataOrigin!=='literal')throw new RangeError('invalid literal stacked line authority')
 return createCartesianStackedLinePaths(chart,cx,cy)
}
