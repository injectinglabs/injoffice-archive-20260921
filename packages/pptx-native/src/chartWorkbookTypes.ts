import type {ChartWorkbookRange} from './chartWorkbookRange.js'
import type {NativeWorkbookV2} from '@injoffice/sheets/browser'
type NativeWorkbookUnsupportedV2=NativeWorkbookV2['unsupported'][number]

/** Source-bound descriptors are issued by the PPTX parser. Cache presence is
 * evidence only: its values are never an input to this resolution boundary. */
export interface ChartWorkbookReference {
 readonly kind:'numRef'|'strRef'
 readonly formula:string
 readonly range:ChartWorkbookRange
 readonly cachePresent:boolean
}
export interface ChartWorkbookBinding {
 readonly relationshipId:string
 readonly part:string
 readonly sha256:string
 readonly byteLength:number
 readonly autoUpdate?:boolean
}
export interface ChartWorkbookResolvedValues {
 readonly dataOrigin:'embedded-workbook'
 readonly formula:string
 readonly kind:'numRef'|'strRef'
 readonly workbookPart:string
 readonly workbookRelationshipId:string
 readonly workbookSHA256:string
 readonly workbookRevision:string
 readonly sheetId:string
 readonly sheetName:string
 readonly sheetPart:string
 readonly addresses:readonly string[]
 readonly values:readonly string[]
 readonly chartCacheIgnored:boolean
 readonly visibility:readonly ChartWorkbookCellVisibility[]
 readonly sourceDiagnostics:readonly NativeWorkbookUnsupportedV2[]
}
/** Values are supplied by XLSX dimension records; optional presence identifies
 * a record, not whether its hidden attribute was explicitly written. */
export interface ChartWorkbookCellVisibility {
 readonly sheetState:'visible'|'hidden'|'veryHidden'
 readonly rowHidden?:boolean
 readonly defaultRowsHidden?:boolean
 readonly columnHidden?:boolean
}
