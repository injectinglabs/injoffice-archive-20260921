import type { createDocxWasmClient } from '@injoffice/docx-wasm'
import type { NativeDocxDocumentV1 } from '../../../packages/docs/src/nativeContract'
import { buildDocxRunMutation, editableDocxRuns } from '../../playground/src/docxRoundTrip'
import { documentStructure, documentParagraphFormatting, docxSelection, documentTableSelection } from './formatting'
import {documentDrawings} from './document-media'
import {imageDimensions} from './raster-image'
import { paragraphAppearance } from './document-style'
type Client = Pick<ReturnType<typeof createDocxWasmClient>, 'apply' | 'extract'>
const paragraphKeys = ['alignment','paragraph_style_id','outline_level','spacing_before_twips','spacing_after_twips','line_spacing','line_rule','indent_left_twips','indent_right_twips','first_line_twips','hanging_twips','numbering']
const runKeys = ['font_family', 'font_size_half_points', 'bold', 'italic', 'underline', 'color']
function simpleParagraph(document: NativeDocxDocumentV1, key: string) {
  const target = editableDocxRuns(document).find(run => run.key === key)
  const index = document.body.blocks.findIndex(block => block.paragraph?.id === target?.paragraphId)
  const paragraph = document.body.blocks[index]?.paragraph
  if (!target || !paragraph || paragraph.runs.length !== 1 || !paragraph.edit_policy.allowed_operations.includes('block.insert_after')) throw new Error('Paragraph breaks currently require a simple body paragraph. Use Insert paragraph below for other supported paragraphs.')
  if (Object.keys(paragraph.properties).some(key => !paragraphKeys.includes(key)) || Object.keys(paragraph.runs[0].properties ?? {}).some(key => !runKeys.includes(key))) throw new Error('This paragraph uses formatting that cannot yet be continued automatically. Insert a separate paragraph below.')
  return { target, index, paragraph }
}
/** All intermediate mutations remain private: either the complete edit returns or the caller keeps its prior bytes/draft. */
export async function replaceParagraphLines(client: Client, source: Uint8Array, document: NativeDocxDocumentV1, key: string, text: string, id: () => string) {
  const selected=docxSelection(document,key),index=document.body.blocks.findIndex(block=>block.paragraph?.id===selected?.paragraph.id)
  if(!selected||index<0||!selected.paragraph.edit_policy.allowed_operations.includes('paragraph.split'))throw new Error('Paragraph breaks are not supported in this text segment yet.')
  const lines=text.replace(/\r\n?/g,'\n').split('\n')
  if(lines.length<2||lines.length>100||text.length>256*1024||text.includes('\t'))throw new Error('Paste up to 100 plain-text paragraphs at a time, without tabs.')
  let bytes=source,model=document
  const joined=lines.join('')
  if(selected.target.text!==joined){bytes=await client.apply(bytes,model,buildDocxRunMutation(model,selected.target,joined,id()));model=await client.extract(bytes)}
  let paragraph=model.body.blocks[index].paragraph!
  let run=paragraph.runs[selected.paragraph.runs.findIndex(run=>run.id===selected.run.id)]
  // Empty-list Enter exits only when the entire paragraph is empty.
  const numbering=paragraphAppearance(model,paragraph).paragraph.numbering
  if(text==='\n'&&paragraph.runs.every(run=>!run.text)&&numbering&&numbering.num_id!=='0'){
    const focus=editableDocxRuns(model).find(target=>target.runId===run.id)!
    bytes=await client.apply(bytes,model,documentParagraphFormatting(model,focus.key,{numbering_num_id:'0',numbering_level:0},id()));model=await client.extract(bytes)
    const next=editableDocxRuns(model).find(target=>target.paragraphId===model.body.blocks[index].paragraph!.id)!
    return {bytes,document:model,key:next.key,text:next.text}
  }
  for(let line=0;line<lines.length-1;line++){
    bytes=await client.apply(bytes,model,{protocol:'injoffice.office.mutations',version:1,format:'docx',mutation_id:id(),expected_revision:model.source.package_sha256,payload:{mutations:[{target_kind:'paragraph',target_id:paragraph.id,expected_xml_sha256:paragraph.anchor.xml_sha256,operation:'paragraph.split',split:{run_id:run.id,offset_utf16:lines[line].length}}]}})
    model=await client.extract(bytes);paragraph=model.body.blocks[index+line+1].paragraph!;run=paragraph.runs[0]
  }
  const focus=editableDocxRuns(model).find(target=>target.runId===run.id&&target.partName===run.anchor.part_name)
  if(!focus||focus.text!==lines.at(-1))throw new Error('The new paragraph did not pass readback.')
  return {bytes,document:model,key:focus.key,text:focus.text}
}

export async function replaceEditableDocumentText(client: Client, source: Uint8Array, document: NativeDocxDocumentV1, search: string, replacement: string, id: () => string) {
  if (!search.length || search.length > 1000 || replacement.length > 10000 || /[\r\n\t]/.test(search + replacement) || search === replacement) throw new Error('Enter different single-line search and replacement text.');
  const targets = editableDocxRuns(document).filter(target => target.text.includes(search));
  if (!targets.length) throw new Error('No matching editable text segments.');
  if (targets.length > 500) throw new Error('Replace up to 500 text segments at a time. Narrow your search.');
  const mutations = targets.map(target => {
    const text = target.text.split(search).join(replacement);
    if (text.length > 256 * 1024) throw new Error('A replacement would make a text segment too large.');
    return buildDocxRunMutation(document, target, text, id()).payload.mutations[0];
  });
  const bytes = await client.apply(source, document, { protocol: 'injoffice.office.mutations', version: 1, format: 'docx', mutation_id: id(), expected_revision: document.source.package_sha256, payload: { mutations } });
  return { bytes, document: await client.extract(bytes) };
}

export async function mergeWithPreviousParagraph(client:Client,source:Uint8Array,document:NativeDocxDocumentV1,key:string,text:string,id:()=>string){
  const selected=docxSelection(document,key),index=document.body.blocks.findIndex(block=>block.paragraph?.id===selected?.paragraph.id)
  if(!selected||index<0)throw new Error('Joining is available only in body paragraphs.')
  if(index===0)return null
  if(selected.paragraph.runs[0].id!==selected.run.id)throw new Error('Place the caret at the start of the paragraph to join it.')
  if(!selected.paragraph.edit_policy.allowed_operations.includes('paragraph.join_previous'))return mergeSimpleParagraph(client,source,document,key,text,id)
  let bytes=source,model=document
  if(text!==selected.target.text){bytes=await client.apply(bytes,model,buildDocxRunMutation(model,selected.target,text,id()));model=await client.extract(bytes)}
  const current=model.body.blocks[index].paragraph!,previous=model.body.blocks[index-1].paragraph!,last=previous.runs.at(-1)!,runIndex=previous.runs.length-1,caret=last.text?.length??0
  bytes=await client.apply(bytes,model,{protocol:'injoffice.office.mutations',version:1,format:'docx',mutation_id:id(),expected_revision:model.source.package_sha256,payload:{mutations:[{target_kind:'paragraph',target_id:current.id,expected_xml_sha256:current.anchor.xml_sha256,operation:'paragraph.join_previous',join:{previous_paragraph_id:previous.id,expected_previous_xml_sha256:previous.anchor.xml_sha256}}]}})
  model=await client.extract(bytes)
  const joined=model.body.blocks[index-1].paragraph,run=joined?.runs[runIndex],focus=run&&editableDocxRuns(model).find(target=>target.runId===run.id&&target.partName===run.anchor.part_name)
  if(!focus||focus.text!==last.text)throw new Error('The joined paragraph did not pass readback.')
  return {bytes,document:model,key:focus.key,text:focus.text,caret}
}

async function mergeSimpleParagraph(client: Client, source: Uint8Array, document: NativeDocxDocumentV1, key: string, text: string, id: () => string) {
  const selected = simpleParagraph(document, key);
  if (selected.index === 0) return null;
  const previous = document.body.blocks[selected.index - 1]?.paragraph;
  if (!previous) throw new Error('A paragraph can only join an adjacent text paragraph.');
  const prior = editableDocxRuns(document).find(target => target.paragraphId === previous.id);
  if (!prior) throw new Error('The previous paragraph is not editable.');
  simpleParagraph(document, prior.key);
  if (JSON.stringify(previous.properties) !== JSON.stringify(selected.paragraph.properties) || JSON.stringify(previous.runs[0].properties ?? {}) !== JSON.stringify(selected.paragraph.runs[0].properties ?? {})) throw new Error('Joining paragraphs with different formatting is not supported yet.');
  const merged = prior.text + text;
  if (merged.length > 256 * 1024 || /[\r\n\t]/.test(merged)) throw new Error('Join plain-text paragraphs up to 256 KB.');
  let bytes = source, model = document;
  if (merged !== prior.text) { bytes = await client.apply(bytes, model, buildDocxRunMutation(model, prior, merged, id())); model = await client.extract(bytes); }
  const current = model.body.blocks[selected.index].paragraph!;
  const currentKey = editableDocxRuns(model).find(target => target.paragraphId === current.id)!.key;
  bytes = await client.apply(bytes, model, documentStructure(model, currentKey, 'block.delete', id())); model = await client.extract(bytes);
  const focus = editableDocxRuns(model).find(target => target.paragraphId === model.body.blocks[selected.index - 1].paragraph!.id)!;
  return { bytes, document: model, key: focus.key, text: focus.text, caret: prior.text.length };
}

export async function insertDocumentTable(client: Client, source: Uint8Array, document: NativeDocxDocumentV1, key: string, rows: number, columns: number, id: () => string) {
  const selected=docxSelection(document,key),index=document.body.blocks.findIndex(block=>block.paragraph?.id===selected?.paragraph.id)
  if (!selected || index<0 || !selected.paragraph.edit_policy.allowed_operations.includes('block.insert_after')) throw new Error('Choose a body paragraph to insert a table below it.')
  if (!Number.isInteger(rows)||rows<1||rows>20||!Number.isInteger(columns)||columns<1||columns>12) throw new Error('Tables support 1–20 rows and 1–12 columns.')
  const target=selected.paragraph
  const bytes=await client.apply(source,document,{protocol:'injoffice.office.mutations',version:1,format:'docx',mutation_id:id(),expected_revision:document.source.package_sha256,payload:{mutations:[{target_kind:'paragraph',target_id:target.id,expected_xml_sha256:target.anchor.xml_sha256,operation:'block.insert_after',table:{rows,columns,column_width_twips:Math.floor(8640/columns)}}]}})
  const model=await client.extract(bytes),table=model.body.blocks[index+1]?.table
  if(!table||table.rows.length!==rows||table.rows.some(row=>row.cells.length!==columns))throw new Error('The inserted table did not pass readback.')
  const paragraph=table.rows[0].cells[0].paragraphs[0],focus=editableDocxRuns(model).find(run=>run.paragraphId===paragraph.id)
  if(!focus)throw new Error('The new table cell is not editable.')
  return {bytes,document:model,key:focus.key,text:focus.text}
}

export async function changeDocumentTable(client:Client,source:Uint8Array,document:NativeDocxDocumentV1,key:string,operation:'block.insert_after'|'block.delete',id:()=>string) {
  const table=documentTableSelection(document,key),index=document.body.blocks.findIndex(block=>block.table?.id===table?.id)
  if(!table||index<0||!table.edit_policy.allowed_operations.includes(operation))throw new Error('This table does not support the selected operation.')
  const target={target_kind:'table' as const,target_id:table.id,expected_xml_sha256:table.anchor.xml_sha256}
  const mutation=operation==='block.insert_after'?{...target,operation,text:''}:{...target,operation}
  const bytes=await client.apply(source,document,{protocol:'injoffice.office.mutations',version:1,format:'docx',mutation_id:id(),expected_revision:document.source.package_sha256,payload:{mutations:[mutation]}})
  const model=await client.extract(bytes),paragraph=model.body.blocks[operation==='block.insert_after'?index+1:Math.max(0,index-1)]?.paragraph
  const targets=editableDocxRuns(model),focus=targets.find(run=>run.paragraphId===paragraph?.id)??targets[0]
  return {bytes,document:model,key:focus?.key??'',text:focus?.text??''}
}

export type TableGridOperation='table.row.insert_after'|'table.row.delete'|'table.column.insert_after'|'table.column.delete'
export async function changeDocumentTableGrid(client:Client,source:Uint8Array,document:NativeDocxDocumentV1,key:string,operation:TableGridOperation,id:()=>string){
  const table=documentTableSelection(document,key),selected=docxSelection(document,key)
  if(!table||!selected||!table.edit_policy.allowed_operations.includes(operation))throw new Error('This table does not support the selected row or column operation.')
  const row=table.rows.findIndex(row=>row.cells.some(cell=>cell.paragraphs.some(paragraph=>paragraph.id===selected.paragraph.id)))
  const column=table.rows[row].cells.findIndex(cell=>cell.paragraphs.some(paragraph=>paragraph.id===selected.paragraph.id))
  const rowOperation=operation.startsWith('table.row'),index=rowOperation?row:column,bodyIndex=document.body.blocks.findIndex(block=>block.table?.id===table.id)
  const bytes=await client.apply(source,document,{protocol:'injoffice.office.mutations',version:1,format:'docx',mutation_id:id(),expected_revision:document.source.package_sha256,payload:{mutations:[{target_kind:'table',target_id:table.id,expected_xml_sha256:table.anchor.xml_sha256,operation,index}]}})
  const model=await client.extract(bytes),next=model.body.blocks[bodyIndex].table!
  const inserted=operation.endsWith('insert_after'),nextRow=Math.min(next.rows.length-1,row+(inserted&&rowOperation?1:0)),nextColumn=Math.min(next.rows[nextRow].cells.length-1,column+(inserted&&!rowOperation?1:0))
  const paragraph=next.rows[nextRow].cells[nextColumn].paragraphs[0],focus=editableDocxRuns(model).find(target=>target.paragraphId===paragraph.id)
  if(!focus)throw new Error('The resulting table cell is not editable.')
  return {bytes,document:model,key:focus.key,text:focus.text}
}

export async function insertDocumentImage(client:Client,source:Uint8Array,document:NativeDocxDocumentV1,key:string,imageBytes:Uint8Array,name:string,id:()=>string) {
  const selected=docxSelection(document,key),index=document.body.blocks.findIndex(block=>block.paragraph?.id===selected?.paragraph.id)
  if(!selected||index<0||!selected.paragraph.edit_policy.allowed_operations.includes('block.insert_after'))throw new Error('Choose a body paragraph to insert an image below it.')
  const dimensions=imageDimensions(imageBytes),scale=Math.min(9525,432*12700/dimensions.width,540*12700/dimensions.height)
  let binary='';for(let offset=0;offset<imageBytes.length;offset+=8192)binary+=String.fromCharCode(...imageBytes.subarray(offset,offset+8192))
  const target=selected.paragraph
  const bytes=await client.apply(source,document,{protocol:'injoffice.office.mutations',version:1,format:'docx',mutation_id:id(),expected_revision:document.source.package_sha256,payload:{mutations:[{target_kind:'paragraph',target_id:target.id,expected_xml_sha256:target.anchor.xml_sha256,operation:'block.insert_after',image:{data_base64:btoa(binary),content_type:dimensions.contentType,width_emu:Math.max(1,Math.round(dimensions.width*scale)),height_emu:Math.max(1,Math.round(dimensions.height*scale)),alt_text:name.replace(/[\r\n\t]/g,' ').slice(0,255)}}]}})
  const model=await client.extract(bytes),picture=model.body.blocks[index+1]?.paragraph?.runs.find(run=>run.drawing)?.drawing,paragraph=model.body.blocks[index+2]?.paragraph
  if(!picture?.media_part||picture.content_type!==dimensions.contentType||picture.width_emu!==Math.max(1,Math.round(dimensions.width*scale))||picture.height_emu!==Math.max(1,Math.round(dimensions.height*scale))||!paragraph)throw new Error('The image insertion did not pass readback.')
  const focus=editableDocxRuns(model).find(run=>run.paragraphId===paragraph.id)
  if(!focus||focus.text!=='')throw new Error('The paragraph after the image is not editable.')
  return {bytes,document:model,key:focus.key,text:focus.text}
}

export async function deleteDocumentImage(client:Client,source:Uint8Array,document:NativeDocxDocumentV1,drawingId:string,id:()=>string){
  const drawing=documentDrawings(document).find(value=>value.id===drawingId),index=document.body.blocks.findIndex(block=>block.paragraph?.runs.some(run=>run.drawing?.id===drawingId))
  if(!drawing||index<0||!drawing.edit_policy.allowed_operations.includes('block.delete'))throw new Error('This image cannot be deleted safely.')
  const bytes=await client.apply(source,document,{protocol:'injoffice.office.mutations',version:1,format:'docx',mutation_id:id(),expected_revision:document.source.package_sha256,payload:{mutations:[{target_kind:'drawing',target_id:drawing.id,expected_xml_sha256:drawing.anchor.xml_sha256,operation:'block.delete'}]}})
  const model=await client.extract(bytes),paragraph=model.body.blocks[index]?.paragraph??model.body.blocks[Math.max(0,index-1)]?.paragraph
  const focus=editableDocxRuns(model).find(run=>run.paragraphId===paragraph?.id)??editableDocxRuns(model)[0]
  return {bytes,document:model,key:focus?.key??'',text:focus?.text??''}
}

export async function replaceDocumentImage(client:Client,source:Uint8Array,document:NativeDocxDocumentV1,drawingId:string,imageBytes:Uint8Array,name:string,id:()=>string){
  const drawing=documentDrawings(document).find(value=>value.id===drawingId),index=document.body.blocks.findIndex(block=>block.paragraph?.runs.some(run=>run.drawing?.id===drawingId))
  if(!drawing||index<0||!drawing.edit_policy.allowed_operations.includes('drawing.replace'))throw new Error('This picture cannot be replaced safely.')
  const dimensions=imageDimensions(imageBytes)
  // Fit within the original picture's box, preserving the new raster ratio.
  const scale=Math.min(drawing.width_emu/dimensions.width,drawing.height_emu/dimensions.height)
  const width=Math.max(1,Math.round(dimensions.width*scale)),height=Math.max(1,Math.round(dimensions.height*scale))
  let binary='';for(let offset=0;offset<imageBytes.length;offset+=8192)binary+=String.fromCharCode(...imageBytes.subarray(offset,offset+8192))
  const bytes=await client.apply(source,document,{protocol:'injoffice.office.mutations',version:1,format:'docx',mutation_id:id(),expected_revision:document.source.package_sha256,payload:{mutations:[{target_kind:'drawing',target_id:drawing.id,expected_xml_sha256:drawing.anchor.xml_sha256,operation:'drawing.replace',image:{data_base64:btoa(binary),content_type:dimensions.contentType,width_emu:width,height_emu:height,alt_text:drawing.alt_text??name.replace(/[\r\n\t]/g,' ').slice(0,255)}}]}})
  const model=await client.extract(bytes),picture=model.body.blocks[index]?.paragraph?.runs.find(run=>run.drawing)?.drawing
  if(model.body.blocks.length!==document.body.blocks.length||!picture?.raster||picture.raster.pixel_width!==dimensions.width||picture.raster.pixel_height!==dimensions.height||picture.width_emu!==width||picture.height_emu!==height)throw new Error('The picture replacement did not pass readback.')
  return {bytes,document:model}
}

export function canInsertDocumentPageBreak(document: NativeDocxDocumentV1, key: string): boolean {
  const selected = docxSelection(document, key)
  if (!selected || selected.run.hyperlink || !selected.paragraph.edit_policy.allowed_operations.includes('page_break.insert')) return false
  const owner = (path: string) => path.slice(0, path.lastIndexOf('/'))
  return selected.paragraph.runs.every(run => run.id === selected.run.id || owner(run.anchor.path) !== owner(selected.run.anchor.path))
}

export async function insertDocumentPageBreak(client: Client, source: Uint8Array, document: NativeDocxDocumentV1, key: string, offset: number, id: () => string) {
  const selected = docxSelection(document, key)
  if (!selected || !canInsertDocumentPageBreak(document, key)) throw new Error('Place the caret in a supported body text run to insert a page break.')
  const index = document.body.blocks.findIndex(block => block.paragraph?.id === selected.paragraph.id)
  if (index < 0 || !Number.isSafeInteger(offset) || offset < 0 || offset > selected.target.text.length) throw new Error('Choose a caret inside body text.')
  const runIndex = selected.paragraph.runs.findIndex(run => run.id === selected.run.id)
  const target = selected.paragraph
  const bytes = await client.apply(source, document, { protocol: 'injoffice.office.mutations', version: 1, format: 'docx', mutation_id: id(), expected_revision: document.source.package_sha256, payload: { mutations: [{ target_kind: 'paragraph', target_id: target.id, expected_xml_sha256: target.anchor.xml_sha256, operation: 'page_break.insert', split: { run_id: selected.run.id, offset_utf16: offset } }] } })
  const model = await client.extract(bytes)
  const paragraph = model.body.blocks[index]?.paragraph
  const after = paragraph?.runs[runIndex + 2]
  const focus = after && editableDocxRuns(model).find(run => run.runId === after.id && run.partName === after.anchor.part_name)
  if (paragraph?.runs[runIndex + 1]?.control !== 'page-break' || !focus || focus.text !== selected.target.text.slice(offset)) throw new Error('The page break did not pass readback.')
  return { bytes, document: model, key: focus.key, text: focus.text }
}
