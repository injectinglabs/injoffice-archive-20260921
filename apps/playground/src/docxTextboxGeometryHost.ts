import type {IncomingMessage,ServerResponse} from 'node:http'
import {compileNativeDocxTextboxShapeV1} from '@injoffice/docs/native-textbox-shape-compiler'

const LIMIT=32*1024*1024
/** Same-origin development adapter over the reusable Node compiler. No file
 * lookup, shell commands, arbitrary URL fetches, or persistent storage. */
export async function handleDocxTextboxGeometryRequest(req:IncomingMessage,res:ServerResponse):Promise<boolean>{
 if(req.url?.split('?')[0]!=='/__injoffice/textbox-geometry')return false
 const send=(status:number,value:unknown)=>{res.writeHead(status,{'Content-Type':'application/json','Cache-Control':'no-store'});res.end(JSON.stringify(value))}
 if(req.method!=='POST'){send(405,{error:'POST required'});return true}
 if(req.headers['content-type']!=='application/json'||!req.headers.origin||new URL(req.headers.origin).host!==req.headers.host){send(403,{error:'Same-origin JSON request required'});return true}
 const declared=Number(req.headers['content-length']??0)
 if(!Number.isSafeInteger(declared)||declared<0||declared>LIMIT){send(413,{error:'Request budget exceeded'});return true}
 try{
  let length=0;const chunks:Buffer[]=[]
  for await(const chunk of req){const b=Buffer.isBuffer(chunk)?chunk:Buffer.from(chunk);length+=b.length;if(length>LIMIT){send(413,{error:'Request budget exceeded'});return true}chunks.push(b)}
  const value=JSON.parse(Buffer.concat(chunks,length).toString('utf8'))
  if(!value||typeof value!=='object'||Object.keys(value).sort().join(',')!=='document,evidence,font_base64,index'||typeof value.font_base64!=='string'||value.font_base64.length>22369624||!value.font_base64.length||value.font_base64.length%4!==0||!/^[A-Za-z0-9+/]*={0,2}$/.test(value.font_base64))throw new TypeError('Invalid request')
  const font=Buffer.from(value.font_base64,'base64');if(font.toString('base64')!==value.font_base64)throw new TypeError('Noncanonical font bytes')
  send(200,compileNativeDocxTextboxShapeV1(value.document,value.evidence,value.index,font))
 }catch{send(400,{error:'Textbox source, font or request could not be qualified'})}
 return true
}
