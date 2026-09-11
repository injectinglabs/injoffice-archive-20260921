import {snapshotNativePlainData} from './nativePlainData.js'
/** Read-only source-derived metadata. Never a workbook mutation envelope. */
export interface NativeWorkbookObjectsV1 {
 protocol: 'injoffice.xlsx.preview-objects'
 version: 1
 package_sha256: string
 tables: NativeTablePreviewV1[]
 charts: NativeChartPreviewV1[]
}
export interface NativeTablePreviewV1 {part:string;sheet_part:string;name:string;ref:string;style:string;header_rows:number;total_rows:number;row_stripes:boolean;column_stripes:boolean;warnings:string[]}
export interface NativeChartPreviewV1 {part:string;type:'col'|'bar'|'unsupported';series:NativeChartSeriesPreviewV1[];warnings:string[]}
export interface NativeChartSeriesPreviewV1 {name:string;values:(number|null)[];labels:string[]}

/** Copy and validate the supplemental projection against the opened package. */
export function decodeNativeWorkbookObjectsV1(input:unknown,packageSHA256:string):NativeWorkbookObjectsV1 {
 input=snapshotNativePlainData(input,{maxDepth:12,maxNodes:200000})
 const fail=():never=>{throw new TypeError('Invalid or stale native workbook object preview')}
 const obj=(value:unknown,keys:string[])=>{if(!value||typeof value!=='object'||Array.isArray(value))return fail();const o=value as Record<string,unknown>;if(Object.keys(o).length!==keys.length||keys.some(k=>!Object.hasOwn(o,k)))return fail();return o}
 let textUnits=0
 const text=(v:unknown,max=4096)=>typeof v==='string'&&v.length<=max&&(textUnits+=v.length)<=4*1024*1024&&!/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(v)?v:fail()
 const part=(v:unknown,empty=false)=>{const s=text(v,1024);if(empty&&s==='')return s;if(!s||s.startsWith('/')||s.includes('\\')||s.split('/').some(p=>!p||p==='.'||p==='..'))return fail();return s}
 const list=(v:unknown,max:number)=>Array.isArray(v)&&v.length<=max?v:fail()
 const warnings=(v:unknown)=>list(v,32).map(x=>text(x))
 const bit=(v:unknown)=>typeof v==='boolean'?v:fail()
 const count=(v:unknown)=>v===0||v===1?v:fail()
 const value=obj(input,['protocol','version','package_sha256','tables','charts'])
 if(value.protocol!=='injoffice.xlsx.preview-objects'||value.version!==1||value.package_sha256!==packageSHA256||!/^sha256:[a-f0-9]{64}$/.test(packageSHA256))return fail()
 const tables=list(value.tables,64).map(v=>{const t=obj(v,['part','sheet_part','name','ref','style','header_rows','total_rows','row_stripes','column_stripes','warnings']);return {part:part(t.part),sheet_part:part(t.sheet_part,true),name:text(t.name),ref:text(t.ref,64),style:text(t.style,256),header_rows:count(t.header_rows),total_rows:count(t.total_rows),row_stripes:bit(t.row_stripes),column_stripes:bit(t.column_stripes),warnings:warnings(t.warnings)}})
 let points=0
 const charts=list(value.charts,64).map((v):NativeChartPreviewV1=>{const c=obj(v,['part','type','series','warnings']);if(c.type!=='col'&&c.type!=='bar'&&c.type!=='unsupported')return fail();const series=list(c.series,32).map(v=>{const s=obj(v,['name','values','labels']);const values=list(s.values,1024).map(v=>{if(++points>65536)return fail();return v===null||typeof v==='number'&&Number.isFinite(v)?v:fail()});return {name:text(s.name),values,labels:list(s.labels,1024).map(x=>text(x,256))}});return {part:part(c.part),type:c.type,series,warnings:warnings(c.warnings)}})
 if(tables.length+charts.length>64||new Set([...tables,...charts].map(v=>v.part)).size!==tables.length+charts.length)return fail()
 return {protocol:'injoffice.xlsx.preview-objects',version:1,package_sha256:packageSHA256,tables,charts}
}

export interface NativeCachedChartMarkV1 {series:number;point:number;value:number;x:number;y:number;width:number;height:number}
/** Normalized chart data marks only: no Office layout/style equivalence claim. */
export function layoutNativeCachedChartV1(chart:NativeChartPreviewV1):{marks:NativeCachedChartMarkV1[];minimum:number;maximum:number;baseline:number}|undefined {
 if(chart.type==='unsupported'||chart.series.length===0||chart.series.length>32||chart.series.some(s=>s.values.length>1024)||chart.series.reduce((sum,s)=>sum+s.values.length,0)>1024)return undefined
 const values=chart.series.flatMap(s=>s.values.filter((v):v is number=>v!==null))
 if(!values.length||values.some(v=>!Number.isFinite(v)))return undefined
 const minimum=Math.min(0,...values),maximum=Math.max(0,...values)
 const extent=maximum-minimum
 if(!Number.isFinite(extent))return undefined
 const scale=extent||1,baseline=maximum/scale
 const count=Math.max(...chart.series.map(s=>s.values.length));if(!count)return undefined
 const marks:NativeCachedChartMarkV1[]=[]
 chart.series.forEach((s,series)=>s.values.forEach((value,point)=>{if(value===null)return;const width=0.8/count/chart.series.length,x=(point+0.1)/count+series*width,y=(maximum-Math.max(0,value))/scale,height=Math.abs(value)/scale;marks.push(chart.type==='col'?{series,point,value,x,y,width,height}:{series,point,value,x:(Math.min(0,value)-minimum)/scale,y:x,width:height,height:width})}))
 return {marks,minimum,maximum,baseline:chart.type==='col'?baseline:-minimum/scale}
}
