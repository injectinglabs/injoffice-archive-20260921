import {preflightWire,canonicalWireSha256} from './nativePagePaintWireV1.js'
import {decodeNativeDocxDocument,type NativeDocxDocumentV1} from './nativeContract.js'
import {decodeNativeDocxResolvedLayout,type NativeDocxResolvedLayoutInputV1} from './nativeResolvedLayout.js'
import {decodeNativeDocxPaginationSettings,type NativeDocxPaginationSettingsV1} from './nativePaginationSettings.js'
import {decodeNativeDocxApproximationEligibilityV1,DOCX_APPROXIMATE_PREVIEW_WARNING,DOCX_APPROXIMATE_LINE_BOX_WARNING,type NativeDocxApproximationEligibilityV1} from './nativeApproximationV1.js'
import {projectNativeDocxAutomaticBordersV1} from './nativeAutomaticBorderPreviewV1.js'
import {DOCX_AUTO_BORDER_WARNING} from './nativeAutomaticBorderEvidenceV1.js'
import {projectNativeDocxAbsentFontSizesV1,DOCX_ABSENT_FONT_SIZE_WARNING,type NativeDocxHostDefaultSizePolicyV1} from './nativeAbsentFontSizeV1.js'
import {qualifyNativeDocxFontDescriptorPreviewV1,DOCX_FONT_DESCRIPTOR_PREVIEW_WARNING,type NativeDocxFontSubstitutionEligibilityV1} from './nativeFontDescriptorPreviewV1.js'
import {DOCX_LEGACY_TABLE_ORIGIN_WARNING,DOCX_TABLE_BORDER_RESERVATION_WARNING} from './nativeLegacyTableOriginV1.js'

export interface NativeDocxFontCompositionV1 {
 source_document:NativeDocxDocumentV1
 source_resolved_layout:NativeDocxResolvedLayoutInputV1
 source_pagination_settings:NativeDocxPaginationSettingsV1
 source_font_inventory_json:string
 font_descriptor_eligibility:NativeDocxFontSubstitutionEligibilityV1
 legacy_eligibility?:NativeDocxApproximationEligibilityV1
 automatic_borders?:true
 font_size_policy?:NativeDocxHostDefaultSizePolicyV1
}
export function qualifyNativeDocxFontCompositionV1(value:unknown){
 if(preflightWire(value,'font composition',100000,10000).length)throw new TypeError('Font composition exceeds structural budget')
 // Conservative UTF-8 upper bound, before cloning or invoking model decoders.
 let bytes=0;const stack:unknown[]=[value]
 while(stack.length){const next=stack.pop();bytes+=32;if(typeof next==='string')bytes+=next.length*3;else if(next&&typeof next==='object')for(const[k,v]of Object.entries(next)){bytes+=k.length*3;stack.push(v)};if(bytes>8*1024*1024)throw new TypeError('Font composition exceeds 8 MiB evidence budget')}
 const v=structuredClone(value) as NativeDocxFontCompositionV1
 if(!v||Object.keys(v).filter(k=>!['legacy_eligibility','automatic_borders','font_size_policy'].includes(k)).sort().join(',')!=='font_descriptor_eligibility,source_document,source_font_inventory_json,source_pagination_settings,source_resolved_layout'||v.automatic_borders!==undefined&&v.automatic_borders!==true)throw new TypeError('Invalid closed font composition policy')
 const d=decodeNativeDocxDocument(v.source_document),r=decodeNativeDocxResolvedLayout(v.source_resolved_layout),s=decodeNativeDocxPaginationSettings(v.source_pagination_settings)
 if(!d.ok||!r.ok||!s.ok||s.value.document_id!==d.value.document_id||s.value.revision!==d.value.revision||s.value.package_sha256!==d.value.source.package_sha256)throw new TypeError('Font composition original source is invalid')
 v.font_descriptor_eligibility=qualifyNativeDocxFontDescriptorPreviewV1(v.font_descriptor_eligibility,d.value,r.value,v.source_font_inventory_json)
 let legacy:NativeDocxApproximationEligibilityV1|undefined
 if(v.legacy_eligibility!==undefined){legacy=decodeNativeDocxApproximationEligibilityV1(v.legacy_eligibility,s.value);if(legacy.status!=='eligible')throw new TypeError('Font composition legacy source is ineligible');v.legacy_eligibility=legacy}
 else if(s.value.profile!=='word-modern-default')throw new TypeError('Legacy font preview requires independently source-qualified settings')
 let document=d.value,resolved=r.value
 const reasons:string[]=[]
 if(v.font_descriptor_eligibility.facts.length)reasons.push(DOCX_FONT_DESCRIPTOR_PREVIEW_WARNING)
 if(legacy){reasons.push(...legacy.reasons,DOCX_APPROXIMATE_PREVIEW_WARNING,DOCX_APPROXIMATE_LINE_BOX_WARNING);if(legacy.legacy_table_origins?.length)reasons.push(DOCX_LEGACY_TABLE_ORIGIN_WARNING);if(document.body.blocks.some(b=>b.table))reasons.push(DOCX_TABLE_BORDER_RESERVATION_WARNING)}
 if(v.automatic_borders){const projection=projectNativeDocxAutomaticBordersV1(document,resolved);if(!projection.facts.length)throw new TypeError('Automatic-border composition requires actual qualified source facts');document=projection.document;resolved=projection.resolved;reasons.push(DOCX_AUTO_BORDER_WARNING)}
 if(v.font_size_policy!==undefined){if(!legacy)throw new TypeError('Absent-size font composition requires source-qualified approximation');const projected=projectNativeDocxAbsentFontSizesV1(document,resolved,legacy.absent_font_sizes??[],v.font_size_policy);if(!projected.applied.length)throw new TypeError('Absent-size composition has no source consumers');resolved=projected.resolved;reasons.push(DOCX_ABSENT_FONT_SIZE_WARNING)}
 return {value:v,document,resolved,settings:s.value,legacy,reasons,descriptors:{eligibility:v.font_descriptor_eligibility,inventoryJSON:v.source_font_inventory_json},sha256:canonicalWireSha256(v)}
}
