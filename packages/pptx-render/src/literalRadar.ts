import type {NativeLiteralRadar} from '@injoffice/pptx-native'
import {createChartRadarGeometry} from './chartRadarGeometry.js'
export type {RadarSeriesVector as LiteralRadarVector} from './chartRadarGeometry.js'
export {RADAR_PREVIEW_POLICY} from './chartRadarGeometry.js'
/** The literal wrapper never upgrades a reference/cache record into literals. */
export function createNativeLiteralRadarPaths(chart:NativeLiteralRadar,cx:number,cy:number){
 if(chart.profile!=='literal-radar-v1'||chart.dataOrigin!=='literal')throw new RangeError('invalid literal radar authority')
 return createChartRadarGeometry(chart,cx,cy)
}

export const RADAR_PREVIEW_DISCLOSURE='Source literal radar uses source-radial-plot-v1: closed standard lines or filled polygons, top-first clockwise categories and explicit radial scales. XML series sequence and reversed source axes are retained. Value-axis spokes paint before standard data and after filled data. Host circular plot fitting and bounded angular rounding apply; PowerPoint margins, radial labels, markers and workbook radar are not reproduced.'
