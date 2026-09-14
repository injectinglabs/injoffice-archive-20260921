import {createHash} from 'node:crypto'
import {readFileSync,writeFileSync} from 'node:fs'
import {resolve} from 'node:path'
const directory=resolve(import.meta.dirname,'../packages/pdf/testdata/fonts')
const provenance=JSON.parse(readFileSync(resolve(directory,'PROVENANCE.json'),'utf8'))
const names=['NotoSansDevanagari-Regular.ttf','NotoSansBengali-Regular.ttf']
const faces=names.map(name=>{
 const bytes=new Uint8Array(readFileSync(resolve(directory,name)))
 if(createHash('sha256').update(bytes).digest('hex')!==provenance.files[name])throw new Error(`Font fixture ${name} hash mismatch`)
 return bytes
})
// TTC1 contains unchanged original tables. Only table-directory offsets become
// absolute collection offsets. No outline data, names, or font features change.
const align=value=>Math.ceil(value/4)*4
const output=new Uint8Array(12+4*faces.length+faces.reduce((sum,face)=>sum+align(face.length),0)),view=new DataView(output.buffer)
view.setUint32(0,0x74746366);view.setUint32(4,0x10000);view.setUint32(8,faces.length)
let offset=12+faces.length*4
faces.forEach((face,index)=>{
 const source=new DataView(face.buffer,face.byteOffset,face.byteLength)
 view.setUint32(12+index*4,offset);output.set(face,offset)
 for(let table=0;table<source.getUint16(4);table++){const record=12+table*16;view.setUint32(offset+record+8,source.getUint32(record+8)+offset)}
 offset+=align(face.length)
})
const filename='NotoSans-Devanagari-Bengali.ttc',path=resolve(directory,filename)
const metadata={generator:'scripts/generate-pdf-font-collection.mjs',format:'TTC 1.0',faces:names,sourceCommit:provenance.commit,license:provenance.license,sha256:createHash('sha256').update(output).digest('hex')}
const record=JSON.stringify(metadata,null,2)+'\n'
if(process.argv.includes('--check')){
 if(!Buffer.from(readFileSync(path)).equals(output)||readFileSync(path+'.json','utf8')!==record)throw new Error('Generated PDF collection fixture is stale')
 console.log('PDF collection fixture and source hashes match')
}else{writeFileSync(path,output);writeFileSync(path+'.json',record)}
