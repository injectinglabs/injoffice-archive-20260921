import {assertResolvedWorkbookChart,type NativeResolvedWorkbookChart} from '@injoffice/pptx-native'
import {createCartesianBarPaths,type CartesianBarVector} from './cartesianBarPaths.js'
import {createCartesianConnectedPaths,type CartesianConnectedVector} from './cartesianConnectedPaths.js'

/** Workbook source stays distinct from literal chart profiles. Shared exact
 * geometry operates only after the native source/value admission boundary. */
export function createNativeWorkbookChartPaths(chart:NativeResolvedWorkbookChart,cx:number,cy:number):readonly (CartesianBarVector|CartesianConnectedVector)[]{
 assertResolvedWorkbookChart(chart)
 const data=chart.data
 if(data.dataOrigin!=='embedded-workbook')throw new TypeError('workbook chart requires embedded source values')
 if(data.profile==='workbook-bar-v1')return createCartesianBarPaths(data,cx,cy)
 return createCartesianConnectedPaths(data,data.profile==='workbook-line-v1'?'line':'scatter',cx,cy)
}
