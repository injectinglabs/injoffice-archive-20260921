// Browser-safe native XLSX contract boundary. The root entry also exports the
// Node-qualified HarfBuzz paint path, which is intentionally not evaluated by
// browser-local extract/apply clients.
export {
  WORKBOOK_MUTATION_PROTOCOL,
  WORKBOOK_MUTATION_VERSION,
  WorkbookMutationValidationError,
  decodeWorkbookMutationBatch,
} from './mutationProtocol.js'

export type {
  RangeRef,
  StyleDelta,
  SupportedWorkbookMutation,
  WorkbookMutationBatchV1,
} from './mutationProtocol.js'

export { CHART_MUTATION_KINDS, CHART_TYPES } from './chartMutationProtocol.js'
export type {
  ChartAnchor,
  ChartDeleteMutation,
  ChartIdentity,
  ChartInsertMutation,
  ChartMutation,
  ChartTypeV1,
  ChartUpdateMutation,
} from './chartMutationProtocol.js'
export { decodeNativeEditableChartsV1 } from './nativeEditableChartsV1.js'
export type { NativeEditableChartV1, XlsxNativeChart } from './nativeEditableChartsV1.js'

export { validateNativeWorkbookV1 } from './nativeValidation.js'
export type { NativeWorkbookV1 } from './nativeContract.generated.js'

export {
  NativeWorkbookV2ValidationError,
  assertNativeWorkbookV2,
  decodeNativeWorkbookV2,
  validateNativeWorkbookV2,
} from './nativeValidationV2.js'

export type {
  NativeWorkbookCellV2,
  NativeWorkbookSheetV2,
  NativeWorkbookV2,
} from './nativeContractV2.generated.js'
export { formatNativeSheetCellDisplayV2 } from './nativeCellDisplayV2.js'
export {projectNativeWorkbookV2} from './nativeRenderModelV2.js'
export {createNativeMaximumDigitWidthAuthorityV2,nativeNormalFontDescentEmV1} from './nativeMaximumDigitWidthV2.js'
export {compileNativeSheetGeometryV2,compileNativeStoredRowSheetGeometryV1,isCompiledNativeStoredRowSheetGeometryV1} from './nativeSheetGeometryV2.js'
export type {NativeSheetGeometryV2,NativeStoredRowSheetGeometryV1,NativeSheetViewportV2} from './nativeSheetGeometryV2.js'
export {selectNativeSheetPrintTitleViewportV1} from './nativeSheetPrintTitleViewportV1.js'
export {compileNativeSheetPagePreviewV1} from './nativeSheetPagePreviewV1.js'
export {layoutNativeDrawingObjectsV1} from './nativeDrawingLayoutV1.js'
export type {NativePositionedDrawingV1} from './nativeDrawingLayoutV1.js'
export type {NativeDrawingObjectV1,NativeDrawingAnchorV1,NativeDrawingMarkerV1} from './nativeDrawingObjectsV1.js'
export {layoutNativeFormControlsV1} from './nativeFormControlLayoutV1.js'
export type {NativePositionedFormControlV1,NativeFormControlCaptionBoxV1} from './nativeFormControlLayoutV1.js'
export type {NativeFormControlV1} from './nativeFormControlObjectsV1.js'
export type {NativeSheetHostPagePolicyV1,NativeSheetPreviewPageV1,NativeSheetPagePreviewV1,NativeSheetPreviewRegionV1,NativeSheetPagePreviewOptionsV1} from './nativeSheetPagePreviewV1.js'
export type { NativeSheetCellDisplayFormatResultV2 } from './nativeCellDisplayV2.js'
export { decodeNativeWorkbookObjectsV1, layoutNativeCachedChartV1 } from './nativeObjectsPreviewV1.js'
export {decodeNativeSheetPrintAreasV1,type NativeSheetPrintAreaV1} from './nativeSheetPrintAreasV1.js'
export {decodeNativeSheetPrintTitlesV1,type NativeSheetPrintTitlesV1,type NativeSheetPrintTitleRangeV1} from './nativeSheetPrintTitlesV1.js'
export {selectNativeSheetPrintAreaV1} from './nativeSheetPrintAreaSelectionV1.js'
export {decodeNativeSheetPrintAreaSetsV1,type NativeSheetPrintAreaSetV1} from './nativeSheetPrintAreaSetsV1.js'
export {selectNativeSheetPrintAreaSetV1,compileNativeSheetPrintAreaSetPreviewV1,type NativeSheetPrintAreaSetPreviewV1} from './nativeSheetPrintAreaSetPreviewV1.js'
export {
  compileNativeSheetPrintPagePreviewV1,
  NATIVE_SHEET_PRINT_PAGE_PREVIEW_V1_DPI,
  NATIVE_SHEET_PRINT_PAGE_PREVIEW_V1_PROTOCOL,
} from './nativeSheetPrintPagePreviewV1.js'
export type {
  NativeSheetPrintPageBandV1,
  NativeSheetPrintPageCssRectV1,
  NativeSheetPrintPagePaintV1,
  NativeSheetPrintPagePreviewOptionsV1,
  NativeSheetPrintPagePreviewV1,
  NativeSheetPrintPageRasterPageV1,
  NativeSheetPrintPageRegionV1,
} from './nativeSheetPrintPagePreviewV1.js'
export {parseNativeSheetHeaderFooterV1} from './nativeSheetHeaderFooterV1.js'
export type {
  NativeSheetHeaderFooterAlignV1,
  NativeSheetHeaderFooterFactsV1,
  NativeSheetHeaderFooterRunV1,
  NativeSheetHeaderFooterSectionV1,
} from './nativeSheetHeaderFooterV1.js'
export { nativeTableFillPreview, nativeTableHeaderTextPreview } from './nativeTableFillPreview.js'
export { nativeTableTotalsTextPreview } from './nativeTableTotalsTextPreview.js'
export { nativeStoredRowPreviewV1 } from './nativeStoredRowsPreviewV1.js'
export type { NativeStoredRowGeometryV1, NativeStoredRowV1 } from './nativeStoredRowsPreviewV1.js'
export { decodeNativeSheetDimensionNeutralityV1, NATIVE_SHEET_DIMENSION_NEUTRALITY_V1_CODES, NATIVE_SHEET_DIMENSION_NEUTRALITY_V1_POLICY, nativeSheetDimensionNeutralityCodesV1 } from './nativeSheetDimensionNeutralityV1.js'
export type { NativeSheetDimensionNeutralityCodeV1, NativeSheetDimensionNeutralityV1 } from './nativeSheetDimensionNeutralityV1.js'
export { nativeTableNumberFormatPreview,formatNativeAccountingTextPreview } from './nativeTableNumberFormatPreview.js'
export { nativeTableBorderPreview } from './nativeTableBorderPreview.js'
export type { NativeTableBorderEdgesV1,NativeTableBorderSideV1 } from './nativeTableBorderPreview.js'
export type { NativeTableNumberDisplayV1 } from './nativeTableNumberFormatPreview.js'
export type { NativeWorkbookObjectsV1, NativeTablePreviewV1, NativeTableNumberFormatV1, NativeTableBorderPreviewV1, NativeChartPreviewV1, NativeChartSeriesPreviewV1, NativeCachedChartMarkV1 } from './nativeObjectsPreviewV1.js'
export { compactNativeGeneralNumberPreviewV1, type NativeCompactGeneralPreviewV1 } from './nativeCompactGeneralPreviewV1.js'

export { decodeNativeConditionalFillPreviewsV1, selectNativeConditionalFillPreviewV1, type NativeConditionalFillPreviewV1, type NativeConditionalFillRuleV1, type NativeConditionalFillCellV1, type NativeConditionalOperatorV1 } from './nativeConditionalFillPreviewV1.js'
export { decodeNativeConditionalScaleFillPreviewsV1, nativeConditionalScaleFillPreview, type NativeConditionalScaleFillPreviewV1, type NativeConditionalScaleFillCellV1, type NativeConditionalScaleFillRangeV1 } from './nativeConditionalScaleFillPreviewV1.js'
export { decodeNativeConditionalBarFillPreviewsV1, nativeConditionalBarFillPreview, type NativeConditionalBarFillPreviewV1, type NativeConditionalBarFillCellV1, type NativeConditionalBarFillRangeV1 } from './nativeConditionalBarFillPreviewV1.js'

export { decodeNativeRichTextPreviewV1, selectNativeRichTextPreviewV1, nativeRichTextRunDisclosureV1, type NativeRichTextPreviewV1, type NativeRichTextCellV1, type NativeRichTextRunV1 } from './nativeRichTextPreviewV1.js'
