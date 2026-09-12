export {
  CSS_PX_PER_INCH,
  PAGE_DIMENSIONS_PX,
  DEFAULT_MARGINS_PX,
  DEFAULT_PAGINATION_OPTIONS,
  contentHeightPx,
  contentWidthPx,
} from './types'
export type { PageSize, Margins, PaginationOptions, BlockBox, PageSlice, PageLayout } from './types'
export { paginate, pageCount, pageOfBlock } from './paginate'
export { renderHeaderFooterTemplate, hasHeaderFooterTokens } from './headerFooterTokens'
export type { HeaderFooterContext } from './headerFooterTokens'
export { groupThreads, openThreadCount, newCommentId, newCommentEntry } from './comments'
export type { CommentEntry, CommentThread } from './comments'
export * from './nativeContract'
export * from './nativePartialContentV1'
export * from './nativePartialNestedTablesV1'
export * from './nativeResolvedLayout'
export * from './nativeNumberingV1'
// Keep approximate implementation seams outside the public package API.
export {
  DOCX_SHAPED_LINES_PROTOCOL,
  DOCX_SHAPED_LINES_VERSION,
  DOCX_SHAPING_REQUEST_PROTOCOL,
  DOCX_SHAPING_REQUEST_VERSION,
  DOCX_SHAPED_LINES_LIMITS,
  type NativeDocxShapingRequestV1,
  type NativeDocxShapingProviders,
  type NativeDocxShapingDiagnosticCode,
  type NativeDocxShapingDiagnosticV1,
  type NativeDocxPositionedGlyphV1,
  type NativeDocxLineFragmentV1,
  type NativeDocxHardBreakV1,
  type NativeDocxShapedLineV1,
  type NativeDocxShapedParagraphV1,
  type NativeDocxShapedListMarkerV1,
  type NativeDocxShapedLinesV1,
  type ShapeNativeDocxLinesResult,
  twipsToMilliPoints,
  halfPointsToMilliPoints,
  type NativeDocxLineIntervalPlanV1,
  shapeNativeDocxLinesV1,
  shapeNativeDocxLinesWithParagraphWidthsV1,
} from './nativeShapingLines'
export type { NativeDocxScriptTransformV1 } from './nativeScriptLayoutV1'
export * from './nativeShapedLinesContract'
export * from './nativePaginationSettings'
export * from './nativePaginationV1'
export * from './nativeSectionColumnsV1'
export * from './nativePaginatedLayoutContract'
// Keep approximate implementation seams outside the public package API.
export {
  DOCX_HEADER_FOOTER_LAYOUT_PROTOCOL,
  DOCX_HEADER_FOOTER_LAYOUT_VERSION,
  type NativeDocxHeaderFooterRegionV1,
  type NativeDocxHeaderFooterDiagnosticCodeV1,
  type NativeDocxHeaderFooterDiagnosticV1,
  type NativeDocxPlacedHeaderFooterLineV1,
  type NativeDocxHeaderFooterPageLayoutV1,
  type NativeDocxHeaderFooterLayoutSuccessV1,
  type NativeDocxHeaderFooterLayoutRefusedV1,
  type NativeDocxHeaderFooterLayoutV1,
  type NativeDocxHeaderFooterLayoutInputV1,
  nativeDocxHeaderFooterLayoutSha256V1,
  layoutNativeDocxHeadersFootersV1,
} from './nativeHeaderFooterLayoutV1'
export * from './nativeImagePagePaintV1'
// Keep approximate implementation seams outside the public package API.
export {
  type NativeDocxPagePaintRequestV1,
  type NativeDocxContentAddressedFaceV1,
  type NativeDocxGlyphOutlineRequestV1,
  type NativeDocxGlyphDesignPathCommandV1,
  type NativeDocxGlyphOutlineResultV1,
  type NativeDocxGlyphOutlineProviderV1,
  type NativeDocxPaintPathCommandV1,
  type NativeDocxFillGlyphPathCommandV1,
  type NativeDocxFillTableCellCommandV1,
  type NativeDocxStrokeTableBorderCommandV1,
  type NativeDocxStrokeNoteSeparatorCommandV1,
  type NativeDocxPaintInlineImageCommandV1,
  type NativeDocxPaintFloatingImageCommandV1,
  type NativeDocxFillTextHighlightCommandV1,
  type NativeDocxStrokeTextUnderlineCommandV1,
  type NativeDocxPagePaintCommandV1,
  type NativeDocxPaintLineV1,
  type NativeDocxPaintPageV1,
  type NativeDocxPagePaintProvenanceV1,
  type NativeDocxPagePaintDiagnosticCode,
  type NativeDocxPagePaintDiagnosticV1,
  type NativeDocxPagePaintSuccessV1,
  type NativeDocxPagePaintRefusedV1,
  type NativeDocxPagePaintV1,
  type DecodeNativeDocxPagePaintRequestV1Result,
  type DecodeNativeDocxPagePaintV1Result,
  type CompileNativeDocxPagePaintV1Result,
  nativeDocxPagePaintFontManifestSha256V1,
  nativeDocxPagePaintShapedLinesSha256V1,
  nativeDocxPagePaintMediaAssetsSha256V1,
  nativeDocxPagePaintPaginatedLayoutSha256V1,
  nativeDocxPagePaintRequestSha256V1,
  nativeDocxPagePaintOutputSha256V1,
  decodeNativeDocxPagePaintRequestV1,
  decodeNativeDocxApproximateComputedPagePaintV1,
  compileNativeDocxPagePaintV1,
  compileNativeDocxApproximatePagePreviewV1,
  compileNativeDocxApproximateComputedPagePreviewV1,
  compileNativeDocxFontSubstitutionPreviewV1,
  decodeNativeDocxPagePaintForRequestV1,
  validateNativeDocxPagePaintForRequestV1,
  DOCX_PAGE_PAINT_REQUEST_PROTOCOL,
  DOCX_PAGE_PAINT_REQUEST_VERSION,
  DOCX_PAGE_PAINT_PROTOCOL,
  DOCX_PAGE_PAINT_VERSION,
  DOCX_PAGE_PAINT_LIMITS,
  DOCX_PAGE_PAINT_REQUEST_V1_BINDING_FIELDS,
  DOCX_PAGE_PAINT_V1_BINDING_FIELDS,
  decodeNativeDocxPagePaintV1,
} from './nativePagePaintV1'
export { DOCX_PAGE_FIELD_LIMITS } from './nativePageFieldsV1'
export type { NativeDocxPageFieldVariantV1 } from './nativePageFieldsV1'
export * from './nativeFontInventoryV1'
export * from './nativePagePaintCompilerV1'
export * from './nativeTablePagePaintV1'
export * from './nativeTransactionAdapterV1'
export * from './nativePartialEquationsV1'
