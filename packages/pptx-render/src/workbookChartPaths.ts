import {createChartRadarGeometry,type RadarSeriesVector} from './chartRadarGeometry.js'
import {createCartesianStackedBarPaths} from './cartesianStackedBarPaths.js'
import {createCartesianStackedLinePaths} from './cartesianStackedLinePaths.js'
import {createCartesianBubblePaths,type LiteralBubbleVector} from './literalBubble.js'
import {assertResolvedWorkbookChart,type NativeResolvedWorkbookChart} from '@injoffice/pptx-native'
import {createCartesianBarPaths,type CartesianBarVector} from './cartesianBarPaths.js'
import {createCartesianConnectedPaths,type CartesianConnectedVector} from './cartesianConnectedPaths.js'

/** Workbook source stays distinct from literal chart profiles. Shared exact
 * geometry operates only after the native source/value admission boundary. */
export function createNativeWorkbookChartPaths(chart:NativeResolvedWorkbookChart,cx:number,cy:number):readonly (CartesianBarVector|CartesianConnectedVector|LiteralBubbleVector|RadarSeriesVector)[]{
 assertResolvedWorkbookChart(chart)
 const data=chart.data
 if(data.dataOrigin!=='embedded-workbook')throw new TypeError('workbook chart requires embedded source values')
 // Authority stays in the admitted workbook envelope; geometry receives only
 // its own fields, never a fabricated literal profile.
 if(data.profile==='workbook-radar-v1'){const {profile,dataOrigin,...geometry}=data;return createChartRadarGeometry(geometry,cx,cy)}
 if(data.profile==='workbook-stacked-bar-v1'){const {profile,dataOrigin,...geometry}=data;return createCartesianStackedBarPaths(geometry,cx,cy)}
 if(data.profile==='workbook-stacked-line-v1'){const {profile,dataOrigin,...geometry}=data;return createCartesianStackedLinePaths(geometry,cx,cy)}
 if(data.profile==='workbook-bubble-v1')return createCartesianBubblePaths(data,cx,cy)
 if(data.profile==='workbook-bar-v1')return createCartesianBarPaths(data,cx,cy)
 return createCartesianConnectedPaths(data,data.profile==='workbook-line-v1'?'line':'scatter',cx,cy)
}
