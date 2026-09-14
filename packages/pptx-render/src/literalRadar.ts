import type {NativeLiteralRadar} from '@injoffice/pptx-native'
import {createChartRadarGeometry} from './chartRadarGeometry.js'
export type {RadarSeriesVector as LiteralRadarVector} from './chartRadarGeometry.js'
export {RADAR_PREVIEW_POLICY} from './chartRadarGeometry.js'
/** The literal wrapper never upgrades a reference/cache record into literals. */
export function createNativeLiteralRadarPaths(chart:NativeLiteralRadar,cx:number,cy:number){
 if(chart.profile!=='literal-radar-v1'||chart.dataOrigin!=='literal')throw new RangeError('invalid literal radar authority')
 return createChartRadarGeometry(chart,cx,cy)
}
