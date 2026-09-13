import {createHash} from 'node:crypto'
import {readFileSync,writeFileSync} from 'node:fs'
import {resolve} from 'node:path'
import {gzipSync} from 'node:zlib'

const directory=resolve(import.meta.dirname,'../go/pptxpatch/presetdata')
const xml=readFileSync(resolve(directory,'preset-shapes.xml'))
if(xml.length!==538970||createHash('sha256').update(xml).digest('hex')!=='4a762444d8d85876881c02a5b1dedf6f73006fcd8acb7b4e393435615b37c780')throw Error('Pinned preset catalog XML mismatch')
// No filename/comment or timestamp; OS=unknown makes the header portable.
const compressed=gzipSync(xml,{level:9});compressed[9]=255
const target=resolve(directory,'preset-shapes.xml.gz')
if(process.argv.includes('--check')){
 if(!readFileSync(target).equals(compressed))throw Error('Compressed preset catalog is stale')
 console.log('Compressed preset catalog matches deterministic generation')
}else{writeFileSync(target,compressed);console.log(`Compressed ${xml.length} preset bytes to ${compressed.length}`)}
