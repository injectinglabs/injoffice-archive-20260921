// Browser-safe native DOCX contract and mutation-envelope boundary. The root
// entry also exports Node-qualified shaping and page-paint providers, which a
// browser-local extract/apply client must not evaluate.
export * from './nativeContract.js'
export * from './nativePartialContentV1.js'
export * from './nativePartialNestedTablesV1.js'
export * from './nativePartialEquationsV1.js'
export {decodeNativeDocxResolvedLayout,type NativeDocxResolvedLayoutInputV1} from './nativeResolvedLayout.js'
export * from './nativeTransactionAdapterV1.js'
