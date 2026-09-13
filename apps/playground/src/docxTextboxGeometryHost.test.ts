import {Readable} from 'node:stream'
import type {IncomingMessage,ServerResponse} from 'node:http'
import {describe,it,expect,vi} from 'vitest'
const compile=vi.hoisted(()=>vi.fn(()=>({status:'supported'})))
vi.mock('@injoffice/docs/native-textbox-shape-compiler',()=>({compileNativeDocxTextboxShapeV1:compile}))
import {handleDocxTextboxGeometryRequest} from './docxTextboxGeometryHost'
async function request(body:string,headers:Record<string,string>={},method='POST'){
 const req=Object.assign(Readable.from([Buffer.from(body)]),{url:'/__injoffice/textbox-geometry',method,headers:{host:'localhost:3333',origin:'http://localhost:3333','content-type':'application/json',...headers}}) as unknown as IncomingMessage
 let status=0,text='';const res={writeHead(code:number){status=code},end(value:string){text=value}} as unknown as ServerResponse
 await handleDocxTextboxGeometryRequest(req,res);return {status,text}
}
describe('textbox helper transport',()=>{
 it('requires same-origin JSON and caps request bytes before parsing',async()=>{
  expect((await request('{}',{origin:'https://other.example'})).status).toBe(403)
  expect((await request('{}',{'content-type':'text/plain'})).status).toBe(403)
  expect((await request('{}',{},'GET')).status).toBe(405)
  expect((await request('{}',{'content-length':String(33*1024*1024)})).status).toBe(413)
 })
 it('admits maximum advertised font payload without regex recursion and rejects malformed padding',async()=>{
  compile.mockClear();const font=Buffer.alloc(16*1024*1024,1).toString('base64'),body=JSON.stringify({document:{},evidence:{},index:0,font_base64:font})
  expect((await request(body)).status).toBe(200);expect(compile).toHaveBeenCalledOnce();expect((compile.mock.calls[0] as unknown[])[3]).toHaveLength(16*1024*1024)
  expect((await request(JSON.stringify({document:{},evidence:{},index:0,font_base64:'AQ==='}))).status).toBe(400)
 })
})
