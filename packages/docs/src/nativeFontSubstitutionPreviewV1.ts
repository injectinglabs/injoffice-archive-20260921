import {sha256} from '@noble/hashes/sha2.js'
import {bytesToHex} from '@noble/hashes/utils.js'
import {EXPLICIT_FONT_POLICY_V1,decodeExplicitFontPolicyV1,canonicalExplicitFontPolicyV1,type ExplicitFontPolicyV1} from '@injoffice/font-metrics/layout'
import {validateFontManifest,type NativeFontManifest} from '@injoffice/font-metrics/layout'
import {canonicalWireSha256} from './nativePagePaintWireV1.js'
import {preflightWire,decodeNativeDocxPagePaintV1,DOCX_PAGE_PAINT_PROTOCOL} from './nativePagePaintWireV1.js'
import type {NativeDocxPagePaintV1} from './nativePagePaintV1.js'
import {decodeNativeDocxFontSubstitutionsV1,type NativeDocxFontSubstitutionV1} from './nativeFontSubstitutionEvidenceV1.js'
export const DOCX_FONT_SUBSTITUTION_PREVIEW_PROTOCOL='injoffice.docx.font-substitution-preview' as const
export const DOCX_FONT_SUBSTITUTION_WARNING='Font substitution preview — layout may differ. Missing source fonts may be replaced only by the explicit operator selections disclosed below for read-only rendering; source font names and original file bytes are unchanged.'
export interface NativeDocxFontSubstitutionPreviewV1 {
 protocol:typeof DOCX_FONT_SUBSTITUTION_PREVIEW_PROTOCOL
 version:1
 fidelity:'approximate'
 read_only:true
 policy:typeof EXPLICIT_FONT_POLICY_V1
 operator_policy:ExplicitFontPolicyV1
 policy_sha256:string
 source:{document_id:string;revision:string;package_sha256:string}
 substitutions:NativeDocxFontSubstitutionV1[]
 selected_font_manifest:NativeFontManifest
 reasons:string[]
 status:NativeDocxPagePaintV1['status']
 pages:NativeDocxPagePaintV1['pages']
 resources:NativeDocxPagePaintV1['resources']
 diagnostics:NativeDocxPagePaintV1['diagnostics']
 rendering_provenance:NativeDocxPagePaintV1['provenance']
}
export const nativeDocxFontPolicySha256V1=(policy:unknown)=>`sha256:${bytesToHex(sha256(new TextEncoder().encode(canonicalExplicitFontPolicyV1(policy))))}`
export function decodeNativeDocxFontSubstitutionPreviewV1(value:unknown):NativeDocxFontSubstitutionPreviewV1{
 if(preflightWire(value,'font substitution preview').length)throw new TypeError('Invalid bounded font preview')
 const v=structuredClone(value) as NativeDocxFontSubstitutionPreviewV1
 if(!v||Object.keys(v).sort().join(',')!=='diagnostics,fidelity,operator_policy,pages,policy,policy_sha256,protocol,read_only,reasons,rendering_provenance,resources,selected_font_manifest,source,status,substitutions,version'||v.protocol!==DOCX_FONT_SUBSTITUTION_PREVIEW_PROTOCOL||v.version!==1||v.fidelity!=='approximate'||v.read_only!==true||v.policy!==EXPLICIT_FONT_POLICY_V1)throw new TypeError('Invalid font preview envelope')
 v.operator_policy=decodeExplicitFontPolicyV1(v.operator_policy)
 if(v.policy_sha256!==nativeDocxFontPolicySha256V1(v.operator_policy))throw new TypeError('Font preview policy hash mismatch')
 v.substitutions=decodeNativeDocxFontSubstitutionsV1(v.substitutions)
 const paint=decodeNativeDocxPagePaintV1({protocol:DOCX_PAGE_PAINT_PROTOCOL,version:1,status:v.status,provenance:v.rendering_provenance,diagnostics:v.diagnostics,pages:v.pages,resources:v.resources})
 if(!paint.ok||!v.source||Object.keys(v.source).sort().join(',')!=='document_id,package_sha256,revision'||v.source.document_id!==v.rendering_provenance.document_id||v.source.revision!==v.rendering_provenance.revision||v.source.package_sha256!==v.rendering_provenance.package_sha256)throw new TypeError('Font preview does not join original source')
 const manifest=validateFontManifest(v.selected_font_manifest)
 if(!manifest.ok||manifest.value.manifestId!==v.rendering_provenance.font_manifest.manifest_id||manifest.value.revision!==v.rendering_provenance.font_manifest.revision||canonicalWireSha256(manifest.value)!==v.rendering_provenance.font_manifest.sha256)throw new TypeError('Font preview manifest does not join rendering provenance')
 for(const r of v.substitutions){
  const face=manifest.value.faces.find(f=>f.faceId===r.face_id)
  if(!face||face.source.kind!=='host'||face.source.contentDigest!==r.font_digest||face.family!==r.selected_family||face.weight!==r.weight||face.style!==r.style||face.stretch!==100)throw new TypeError('Substitution record does not join selected font manifest')
 }
 if(!Array.isArray(v.reasons)||v.reasons.length>10001||!v.reasons.includes(DOCX_FONT_SUBSTITUTION_WARNING)||v.reasons.some(r=>typeof r!=='string'||r.length>2048))throw new TypeError('Font preview lacks persistent warning')
 for(const r of v.substitutions)if(!v.operator_policy.mappings.some(m=>m.sourceFamily.toLowerCase()===r.source_family.toLowerCase()&&m.targetFamily===r.selected_family&&m.weight===r.weight&&m.style===r.style))throw new TypeError('Font preview record lacks explicit operator mapping')
 return v
}
