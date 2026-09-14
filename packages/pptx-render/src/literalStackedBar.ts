import {createCartesianStackedBarPaths,type CartesianStackedBarInput} from './cartesianStackedBarPaths.js'
export type {CartesianBarVector as LiteralStackedBarVector} from './cartesianBarPaths.js'
export type LiteralStackedBarInput=CartesianStackedBarInput&{readonly profile:'literal-stacked-bar-v1';readonly dataOrigin:'literal'}
/** Exact literal profile only. Native attachment admission separately verifies
 * source ownership and label metadata; this helper produces source-colored paths. */
export function createNativeLiteralStackedBarPaths(chart:LiteralStackedBarInput,cx:number,cy:number){
 if(chart.profile!=='literal-stacked-bar-v1'||chart.dataOrigin!=='literal')throw new RangeError('invalid literal stacked bar authority')
 return createCartesianStackedBarPaths(chart,cx,cy)
}
