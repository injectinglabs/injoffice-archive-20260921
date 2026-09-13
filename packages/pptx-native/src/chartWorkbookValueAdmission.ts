import type {ChartWorkbookResolvedValues} from './chartWorkbookTypes.js'
const admitted=new WeakSet<ChartWorkbookResolvedValues>()
export function admitChartWorkbookValues(value:ChartWorkbookResolvedValues):ChartWorkbookResolvedValues{admitted.add(value);return value}
export function assertChartWorkbookValues(value:ChartWorkbookResolvedValues):void{if(!admitted.has(value))throw new TypeError('workbook values must come from the validated source resolver')}
