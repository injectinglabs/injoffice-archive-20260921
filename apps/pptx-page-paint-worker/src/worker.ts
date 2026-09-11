import {compilePptxPreview} from './compile.js'
const MAX=16*1024*1024
let pending=Buffer.alloc(0)
process.stdin.on('data',(chunk:Buffer)=>{if(pending.length+chunk.length>MAX+4)process.exit(1);pending=Buffer.concat([pending,chunk])})
process.stdin.on('end',()=>{void(async()=>{
 if(pending.length<4||pending.readUInt32BE(0)!==pending.length-4||pending.length===4)throw new Error('Invalid worker frame')
 const request=JSON.parse(pending.subarray(4).toString('utf8'))
 if(request.protocol!=='injoffice.pptx.preview-worker'||request.version!==1||request.id!=='preview'||request.op!=='render')throw new Error('Invalid worker envelope')
 let response:unknown
 try{response={protocol:request.protocol,version:1,id:request.id,ok:true,result:await compilePptxPreview(request.input)}}catch(error){response={protocol:request.protocol,version:1,id:request.id,ok:false,error:{message:error instanceof Error?error.message.slice(0,1024):'Native preview failed'}}}
 const output=Buffer.from(JSON.stringify(response));if(output.length>MAX)throw new Error('Preview response exceeded byte budget')
 const header=Buffer.alloc(4);header.writeUInt32BE(output.length);process.stdout.write(Buffer.concat([header,output]))
})().catch(()=>{process.exitCode=1})})
