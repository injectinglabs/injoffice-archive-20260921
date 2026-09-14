import type {NativeLiteralBarAxis} from './types.js'
import type {ChartWorkbookBinding,ChartWorkbookReference} from './chartWorkbookTypes.js'

export interface NativePptxWorkbookChartSeries {
 readonly index:number;readonly order:number;readonly title?:string
 readonly titleReference?:ChartWorkbookReference;readonly categoryReference?:ChartWorkbookReference
 readonly sizeReference?:ChartWorkbookReference;readonly xReference?:ChartWorkbookReference;readonly valueReference:ChartWorkbookReference
 readonly colors?:readonly string[];readonly color?:string;readonly widthEmu?:number
}
export interface NativePptxWorkbookChartSource {
 readonly family:'bar'|'line'|'scatter'|'bubble';readonly bubbleScale?:number;readonly sizeRepresents?:'area'|'w';readonly barDirection?:'col'|'bar';readonly gapWidth?:number
 readonly xAxis:NativeLiteralBarAxis;readonly yAxis:NativeLiteralBarAxis
 readonly series:readonly NativePptxWorkbookChartSeries[];readonly plotVisibleOnly:false
 readonly dispBlanksAs?:'gap'|'zero'|'span'
}
export interface NativePptxInspectedWorkbookChart {
 readonly slide_id:string;readonly slide_index:number;readonly slide_part:string;readonly slide_sha256:string
 readonly element_id:string;readonly object_id:string
 /** SHA-256 of the exact source p:graphicFrame XML bytes, not a descriptor hash. */
 readonly frame_sha256:string
 readonly chart_relationship_id:string;readonly chart_part:string;readonly chart_sha256:string
 readonly source:NativePptxWorkbookChartSource;readonly workbook:ChartWorkbookBinding
}
export interface NativePptxInspectedChartWorkbook {
 readonly part:string;readonly sha256:string;readonly byteLength:number;readonly bytesBase64:string
}
export interface NativePptxChartWorkbookInspection {
 readonly protocol:'pptx-chart-workbook-inspection-v1';readonly package_sha256:string;readonly source_revision:string
 readonly charts:readonly NativePptxInspectedWorkbookChart[]
 readonly workbooks:readonly NativePptxInspectedChartWorkbook[]
 readonly omissions:readonly {readonly slide_id:string;readonly object_id:string;readonly reason:string}[]
}
