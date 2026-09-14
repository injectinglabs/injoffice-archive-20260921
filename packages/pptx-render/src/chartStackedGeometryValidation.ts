const integer=(value:unknown,min:number,max:number)=>typeof value==='number'&&Number.isInteger(value)&&!Object.is(value,-0)&&value>=min&&value<=max
const rgb=(value:unknown)=>typeof value==='string'&&value.length===7&&/^#[0-9A-F]{6}$/.test(value)
function record(value:unknown,required:string,optional=''):Record<string,unknown>{
 if(value===null||typeof value!=='object'||Array.isArray(value))throw new RangeError('invalid stacked chart record')
 const keys=required.split(' '),allowed=new Set([...keys,...optional.split(' ').filter(Boolean)])
 if(keys.some(key=>!Object.hasOwn(value,key))||Object.keys(value).some(key=>!allowed.has(key)))throw new RangeError('invalid stacked chart record fields')
 return value as Record<string,unknown>
}
function dense(value:unknown,min:number,max:number):unknown[]{
 if(!Array.isArray(value)||value.length<min||value.length>max)throw new RangeError('invalid stacked chart array bounds')
 for(let i=0;i<value.length;i++)if(!Object.hasOwn(value,i))throw new RangeError('sparse stacked chart array')
 return value
}
function axis(value:unknown):void{
 const a=record(value,'id crossAxisId orientation position deleted','color widthEmu min max crossesAt labels')
 if(!integer(a.id,0,4294967295)||!integer(a.crossAxisId,0,4294967295)||!['minMax','maxMin'].includes(a.orientation as string)||!['b','l'].includes(a.position as string)||typeof a.deleted!=='boolean')throw new RangeError('invalid stacked chart axis identity')
 if(a.deleted?(a.color!==undefined||a.widthEmu!==undefined):(!rgb(a.color)||!integer(a.widthEmu,1,20116800)))throw new RangeError('invalid stacked chart axis paint')
 // Axis-label layout is admitted by the separate native attachment validator;
 // this helper paints only axis lines. Retain only its existing metadata slot.
}
/** Strict preflight supplements legacy geometry without changing its admission.
 * The following scaffold still checks axis pairing, decimal scales and cross0. */
export function validateStackedChartGeometry(value:unknown,family:'bar'|'line'):void{
 if(family!=='bar'&&family!=='line')throw new RangeError('invalid stacked chart family')
 const chart=record(value,family==='bar'?'grouping overlap barDirection gapWidth categories series categoryAxis valueAxis':'grouping categories series xAxis yAxis','profile dataOrigin')
 if(!['stacked','percentStacked'].includes(chart.grouping as string)||(chart.profile!==undefined&&chart.profile!==`literal-stacked-${family}-v1`)||(chart.dataOrigin!==undefined&&chart.dataOrigin!=='literal'))throw new RangeError('invalid stacked chart profile')
 const categories=dense(chart.categories,1,256)
 if(categories.some(c=>typeof c!=='string')||categories.reduce<number>((n,c)=>n+(c as string).length,0)>32768)throw new RangeError('invalid stacked chart categories')
 const series=dense(chart.series,1,16)
 for(const raw of series){
  const s=record(raw,family==='bar'?'index order values colors':'index order values color widthEmu','title')
  if(!integer(s.index,0,4294967295)||!integer(s.order,0,15)||(s.title!==undefined&&(typeof s.title!=='string'||s.title.length>1024)))throw new RangeError('invalid stacked chart series identity')
  const values=dense(s.values,categories.length,categories.length)
  if(values.some(v=>typeof v!=='string'))throw new RangeError('invalid stacked chart values')
  if(family==='bar'){if(dense(s.colors,categories.length,categories.length).some(color=>!rgb(color)))throw new RangeError('invalid stacked chart fill')}
  else if(!rgb(s.color)||!integer(s.widthEmu,1,20116800))throw new RangeError('invalid stacked chart stroke')
 }
 if(family==='bar'){
  if(!['column','bar'].includes(chart.barDirection as string)||chart.overlap!==100||!integer(chart.gapWidth,0,500))throw new RangeError('invalid stacked bar geometry')
  axis(chart.categoryAxis);axis(chart.valueAxis)
 }else{axis(chart.xAxis);axis(chart.yAxis)}
}
