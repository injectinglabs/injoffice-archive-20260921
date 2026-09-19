const {Worker}=require('node:worker_threads');
const {createHash}=require('node:crypto');
const path=require('node:path');
const {validateBytes}=require('./file-store.cjs');
let active;
function exportDocxPdf(input,{workerPath=path.join(__dirname,'docx-pdf-worker.mjs'),onProgress=()=>{},timeoutMs=60000}={}){
 if(active)return Promise.reject(new Error('Another document PDF export is running.'));
 if(!input||typeof input.requestId!=='string'||!/^[a-zA-Z0-9-]{1,80}$/.test(input.requestId))return Promise.reject(new Error('Invalid PDF export request.'));
 if(input.mode!==undefined&&!['original','preview'].includes(input.mode))return Promise.reject(new Error('Invalid PDF font handling mode.'));
 const bytes=validateBytes(input.bytes);
 if(bytes.length>32*1024*1024)return Promise.reject(new Error('Document PDF export supports files up to 32 MiB.'));
 if(typeof input.revision!=='string'||input.revision!==`sha256:${createHash('sha256').update(bytes).digest('hex')}`)return Promise.reject(new Error('The document revision changed. Try exporting again.'));
 return new Promise((resolve,reject)=>{
  let worker;try{worker=new Worker(workerPath,{workerData:{bytes,revision:input.revision,mode:input.mode??'original'},resourceLimits:{maxOldGenerationSizeMb:512}})}catch(error){reject(error);return}
  let settled=false;
  const finish=(error,result)=>{if(settled)return;settled=true;clearTimeout(timer);active=undefined;void worker.terminate();error?reject(error):resolve(result)};
  const timer=setTimeout(()=>finish(new Error('PDF export exceeded 60 seconds. The document is unchanged.')),timeoutMs);
  active={requestId:input.requestId,cancel:()=>finish(new Error('PDF export canceled.'))};
  worker.on('message',result=>{
   if(result?.progress){if(['reading','layout','outlining','painting'].includes(result.progress))onProgress(result.progress);return}
   if(result?.error){finish(new Error(String(result.error).slice(0,4096)));return}
   if(!(result?.bytes instanceof Uint8Array)||result.bytes.length>64*1024*1024||!Buffer.from(result.bytes.subarray(0,5)).equals(Buffer.from('%PDF-'))){finish(new Error('The PDF exporter returned invalid output.'));return}
   finish(null,result.bytes);
  });
  worker.on('error',error=>finish(error));
  worker.on('exit',()=>finish(new Error('The PDF exporter stopped before completing.')));
 });
}
function cancelDocxPdf(requestId){if(active?.requestId!==requestId)return false;active.cancel();return true}
module.exports={exportDocxPdf,cancelDocxPdf};
