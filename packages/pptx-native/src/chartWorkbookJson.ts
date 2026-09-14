/** Bounded inspection transport scanner. JSON.parse alone loses duplicate keys
 * and unsafe integer lexemes before the source-identity validator can see them. */
export function parseChartWorkbookJson(source:string):unknown {
 if(typeof source!=='string'||source.length>33554432||new TextEncoder().encode(source).byteLength>33554432)throw new RangeError('chart workbook JSON byte budget exceeded')
 let at=0,nodes=0
 const white=()=>{while(at<source.length&&' \n\r\t'.includes(source[at]!))at++}
 const fail=():never=>{throw new TypeError('invalid chart workbook JSON')}
 const string=():string=>{
  const start=at;if(source[at++]!=='"')return fail()
  while(at<source.length){const ch=source[at++]!;if(ch==='"'){const value:unknown=JSON.parse(source.slice(start,at));if(typeof value!=='string'||/[\uD800-\uDFFF]/u.test(value))return fail();return value}if(ch==='\\'){if(at>=source.length)return fail();at++}}
  return fail()
 }
 const value=(depth:number):void=>{
  if(depth>32||++nodes>100000)throw new RangeError('chart workbook JSON structural budget exceeded')
  white();const ch=source[at]
  if(ch==='"'){string();return}
  if(ch==='{'||ch==='['){
   at++;white();const object=ch==='{',end=object?'}':']',keys=new Set<string>()
   if(source[at]===end){at++;return}
   while(at<source.length){
    if(object){white();const key=string();if(keys.has(key))throw new TypeError('duplicate chart workbook JSON key');keys.add(key);white();if(source[at++]!==':')return fail()}
    value(depth+1);white();if(source[at]===end){at++;return}if(source[at++]!==',')return fail()
   };return fail()
  }
  for(const literal of ['true','false','null'])if(source.startsWith(literal,at)){at+=literal.length;return}
  const start=at;while(at<source.length&&!',[}] \n\r\t'.includes(source[at]!))at++
  const raw=source.slice(start,at),number=Number(raw)
  if(!/^-?(?:0|[1-9][0-9]*)$/.test(raw)||!Number.isSafeInteger(number)||Object.is(number,-0))return fail()
 }
 value(0);white();if(at!==source.length)return fail();return JSON.parse(source)
}
