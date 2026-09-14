/** Joint body/footnote flow. Planning is source-bound and commits no page output. */
import {decodeNativeDocxPaginationRequestV1, DOCX_PAGINATION_LIMITS, type NativeDocxPaginationRequestV1} from './nativePaginationV1.js'
import {DOCX_NOTE_PAGINATION_LIMITS, type NativeDocxNotePaginationRefusalV1} from './nativeNotePaginationV1.js'
import {qualifyNativeDocxSectionColumnsV1, type NativeDocxQualifiedSectionGeometryV1} from './nativeSectionColumnsV1.js'
import type {NativeDocxStoryV1} from './nativeContract.js'
import type {NativeDocxShapedLineV1} from './nativeShapingLines.js'

interface Slice {paragraph_id:string;start:number;end:number}
interface NoteSlice {story_id:string;start:number;end:number}
interface NoteLine {paragraph_id:string;line:NativeDocxShapedLineV1;paragraph_start:number;paragraph_end:number;keep:boolean;widow:boolean}
export interface NativeDocxFootnoteFlowV1 {
 status:'planned'
 section_id:string
 geometry:NativeDocxQualifiedSectionGeometryV1
 pages:{body:Slice[];notes:NoteSlice[];separator?:NativeDocxStoryV1;note_height:number}[]
 notes:Map<string,{story:NativeDocxStoryV1;reference_run_id:string;number:number;lines:NoteLine[];heights:number[]}>
}

/** Undefined leaves other established profiles in charge. A qualified but
 * impossible flow returns one refusal, never a partial plan. */
export function planNativeDocxFootnoteFlowV1(value:unknown):NativeDocxFootnoteFlowV1|NativeDocxNotePaginationRefusalV1|undefined {
 const decoded=decodeNativeDocxPaginationRequestV1(value)
 if(!decoded.ok)return undefined
 const request:NativeDocxPaginationRequestV1=decoded.value
 const {document,resolved_layout:resolved,shaped_lines:shaped,pagination_settings:settings}=request
 if(request.column_shaped_lines||settings.profile!=='word-modern-default'||settings.diagnostics.length||settings.even_and_odd_headers||settings.mirror_margins||settings.gutter_at_top||
  document.sections.length!==1||document.headers.length||document.footers.length||document.comments.length||document.comment_stories.length||resolved.tables.length||resolved.numbering_source)return undefined
 const section=document.sections[0]!
 if(section.page.columns!==1||section.header_refs.length||section.footer_refs.length||section.title_page||section.break_type!=='next-page')return undefined
 const geometry=qualifyNativeDocxSectionColumnsV1(section)
 if(!geometry.ok)return undefined
 const column=geometry.value.columns[0]!,height=column.height_millipoints
 if(shaped.available_width_millipoints!==column.width_millipoints||section.starts_at_block_id!==document.body.blocks[0]?.id)return undefined
 const content=document.notes.filter(n=>(n.note_role??'content')==='content')
 if(!content.length||document.notes.some(n=>n.kind!=='footnote'))return undefined
 const fail=(message:string,scope_id=document.document_id,code:NativeDocxNotePaginationRefusalV1['code']='note-overflow-unsupported'):NativeDocxNotePaginationRefusalV1=>({scope_id,code,message})
 if(content.length>DOCX_NOTE_PAGINATION_LIMITS.maxNotes)return fail('Footnote count exceeds the flow budget',document.document_id,'resource-limit')
 const shapes=new Map(shaped.paragraphs.map(p=>[p.paragraph_id,p]))
 const properties=new Map(resolved.paragraphs.map(p=>[p.paragraph_id,p]))
 const body=[]
 for(const block of document.body.blocks){if(block.kind!=='paragraph'||!block.paragraph)return undefined;body.push(block.paragraph)}
 if(!body.length)return undefined
 const bodyIDs=new Set(body.map(p=>p.id))
 let lineCount=0,totalHeight=0
 for(const paragraph of [...body,...content.flatMap(n=>n.blocks.map(b=>b.paragraph))]){
  if(!paragraph)return undefined
  const shape=shapes.get(paragraph.id),resolvedParagraph=properties.get(paragraph.id)
  if(!shape||!resolvedParagraph||paragraph.properties.numbering||resolvedParagraph.numbering||resolvedParagraph.properties.keep_next||resolvedParagraph.properties.bidi||
   shape.spacing_before_millipoints||shape.spacing_after_millipoints||shape.direction!=='ltr'||!shape.lines.length||
   shape.block_advance_millipoints!==shape.lines.reduce((sum,l)=>sum+l.line_height_millipoints,0)||
   shape.lines.some(l=>l.hard_break_after||l.exclusion_start_millipoints!==undefined)||
   paragraph.runs.some(r=>r.page_field||r.layout_page_field||r.drawing||r.control))return undefined
  if(!bodyIDs.has(paragraph.id)&&resolvedParagraph.properties.page_break_before)return undefined
  if(shape.lines.some(l=>l.line_height_millipoints>height)){
   if(bodyIDs.has(paragraph.id))return undefined
   return fail('One shaped note line exceeds its page column',paragraph.id,'note-structure-unsupported')
  }
  totalHeight+=shape.block_advance_millipoints
  if(totalHeight>height*DOCX_PAGINATION_LIMITS.maxPages)return fail('Body and footnote height exceeds the page budget',document.document_id,'resource-limit')
  lineCount+=shape.lines.length
 }
 if(lineCount>DOCX_NOTE_PAGINATION_LIMITS.maxNoteLines)return fail('Body and footnote lines exceed the flow budget',document.document_id,'resource-limit')
 const notes:NativeDocxFootnoteFlowV1['notes']=new Map()
 const sourceNotes=new Map(content.map(n=>[n.id,n]))
 const referencesByLine=new Map<string,string[]>()
 const referenceIDs=new Set(body.flatMap(p=>p.runs.filter(r=>r.reference).map(r=>r.id)))
 const referenceLines=new Map<string,NativeDocxShapedLineV1[]>()
 for(const p of body)for(const line of shapes.get(p.id)!.lines)for(const id of new Set(line.fragments.map(f=>f.source_id))){
  if(referenceIDs.has(id)){const found=referenceLines.get(id);if(!found)referenceLines.set(id,[line]);else if(found.length<2)found.push(line)}
 }
 for(const paragraph of body)for(const run of paragraph.runs){
  const reference=run.reference
  if(!reference)continue
  if(reference.kind!=='footnote'||reference.role==='label')return undefined
  const story=sourceNotes.get(reference.target_id)
  const lines=referenceLines.get(run.id)??[]
  if(!story||notes.has(story.id)||lines.length!==1)return fail('Each footnote requires one uniquely shaped body reference',run.id,'note-reference-ambiguous')
  const entries:NoteLine[]=[]
  for(const block of story.blocks){
   if(!block.paragraph)return undefined
   const shape=shapes.get(block.id)!,p=properties.get(block.id)!.properties,start=entries.length,end=start+shape.lines.length
   for(const line of shape.lines)entries.push({paragraph_id:block.id,line,paragraph_start:start,paragraph_end:end,keep:p.keep_lines===true,widow:p.widow_control??true})
  }
  const heights=[0]
  for(const entry of entries)heights.push(heights.at(-1)!+entry.line.line_height_millipoints)
  if(!Number.isSafeInteger(heights.at(-1)))return fail('Footnote height exceeds safe integer geometry',story.id,'resource-limit')
  notes.set(story.id,{story,reference_run_id:run.id,number:notes.size+1,lines:entries,heights})
  const lineID=lines[0]!.id
  const ids=referencesByLine.get(lineID)??[];ids.push(story.id);referencesByLine.set(lineID,ids)
 }
 if(notes.size!==content.length)return fail('Every footnote must be referenced exactly once',document.document_id,'note-reference-ambiguous')
 const ordinary=document.notes.filter(n=>n.note_role==='separator'),continuations=document.notes.filter(n=>n.note_role==='continuation-separator')
 if(ordinary.length!==1)return fail('Footnote flow requires one ordinary separator',document.document_id,'note-separator-unsupported')
 function separator(continued:boolean):{story:NativeDocxStoryV1;height:number}|NativeDocxNotePaginationRefusalV1 {
  const candidates=continued?continuations:ordinary
  const story=candidates[0]
  if(candidates.length!==1||!story||story.relationship_id!==ordinary[0]!.relationship_id||story.part_name!==ordinary[0]!.part_name||story.blocks.length!==1||!story.blocks[0]?.paragraph||story.blocks[0].paragraph.runs.length)return fail('Footnote flow requires one exact source-bound separator',story?.id, 'note-separator-unsupported')
  const scopes=new Set([story.id,story.blocks[0].id])
  if(document.unsupported.some(d=>scopes.has(d.scope_id))||resolved.diagnostics.some(d=>scopes.has(d.scope_id))||shaped.diagnostics.some(d=>scopes.has(d.scope_id)||d.source_id!==undefined&&scopes.has(d.source_id)))return fail('Activated footnote separator has unsupported semantics',story.id,'note-separator-unsupported')
  const shape=shapes.get(story.blocks[0].id)
  if(!shape&&continued)return fail('Footnote continuation requires activated separator shaping',story.id,'note-continuation-shaping-required')
  if(!shape||shape.lines.length!==1||shape.lines[0]!.fragments.length||shape.spacing_before_millipoints||shape.spacing_after_millipoints)return fail('Footnote separators require one zero-spacing instruction line',story.id,'note-separator-unsupported')
  return {story,height:shape.lines[0]!.line_height_millipoints}
 }
 const rule=separator(false)
 if('code' in rule)return rule
 type Pending={story_id:string;start:number}
 // Consume notes in source order. Only the last selected note may be partial;
 // a newly activated reference must retain a legal first slice on this page.
 function select(queue:Pending[],capacity:number):{slices:NoteSlice[];height:number} {
  const slices:NoteSlice[]=[];let used=0
  for(const pending of queue){
   const note=notes.get(pending.story_id)!,entries=note.lines,heights=note.heights,start=pending.start
   let low=start,high=entries.length
   while(low<high){const mid=Math.ceil((low+high)/2);if(used+(heights[mid]!-heights[start]!)<=capacity)low=mid;else high=mid-1}
   let end=low
   while(end>start&&end<entries.length){
    const last=entries[end-1]!,first=Math.max(start,last.paragraph_start)
    if(entries[end]!.paragraph_id!==last.paragraph_id)break
    if(last.keep||last.widow&&end-first<2)end=first
    else if(last.widow&&last.paragraph_end-end<2)end=last.paragraph_end-2
    else break
   }
   if(end===start)break
   slices.push({story_id:pending.story_id,start,end});used+=heights[end]!-heights[start]!
   if(end<entries.length)break
  }
  return {slices,height:used}
 }
 let continuedRule:ReturnType<typeof separator>|undefined
 const pages:NativeDocxFootnoteFlowV1['pages']=[]
 let pending:Pending[]=[],paragraphIndex=0,lineIndex=0,placedLines=0
 while(paragraphIndex<body.length||pending.length){
  if(pages.length>=DOCX_PAGINATION_LIMITS.maxPages)return fail('Footnote flow exceeds the page budget',document.document_id,'resource-limit')
  const carried=pending.length>0,activeRule=carried?(continuedRule??=separator(true)):rule
  if('code' in activeRule)return activeRule
  let queue=[...pending],bodyHeight=0,selected={slices:[] as NoteSlice[],height:0}
  const remainingNoteHeight=(p:Pending)=>{const n=notes.get(p.story_id)!;return n.heights.at(-1)!-n.heights[p.start]!}
  let queueHeight=queue.reduce((sum,p)=>sum+remainingNoteHeight(p),0)
  const bodySlices:Slice[]=[]
  while(paragraphIndex<body.length){
   const paragraph=body[paragraphIndex]!,shape=shapes.get(paragraph.id)!,props=properties.get(paragraph.id)!.properties
   if(lineIndex===0&&props.page_break_before&&bodySlices.length)break
   const prior=bodySlices.at(-1),onPage=prior?.paragraph_id===paragraph.id
   const remaining=shape.lines.length-lineIndex,widow=props.widow_control??true
   const count=props.keep_lines?remaining:!widow?1:!onPage?remaining<=3?remaining:2:remaining<=2?remaining:1
   const end=lineIndex+count,lines=shape.lines.slice(lineIndex,end),groupHeight=lines.reduce((sum,l)=>sum+l.line_height_millipoints,0)
   const newIDs=lines.flatMap(l=>referencesByLine.get(l.id)??[])
   const incoming=newIDs.map(story_id=>({story_id,start:0}))
   const hasNotes=queue.length+incoming.length>0
   const available=height-bodyHeight-groupHeight-(hasNotes?activeRule.height:0)
   // A new note can start only after every older note finishes. Reuse the
   // already selected prefix instead of copying/re-measuring it per reference.
   const additions=incoming.length?select(incoming,available-queueHeight):undefined
   const candidate=bodySlices.length===0&&!additions?select(queue,available):selected
   const candidateHeight=additions?queueHeight+additions.height:candidate.height
   const fits=available>=candidateHeight&&(!hasNotes||(additions?additions.slices.length===incoming.length:candidate.slices.length>0))
   if(!fits)break
   if(additions){
    if(queue.length){
     if(!selected.slices.length)selected.slices=queue.map(p=>({...p,end:notes.get(p.story_id)!.lines.length}))
     else selected.slices.at(-1)!.end=notes.get(queue.at(-1)!.story_id)!.lines.length
    }
    selected.slices.push(...additions.slices);selected.height=candidateHeight
    queue.push(...incoming);queueHeight+=incoming.reduce((sum,p)=>sum+remainingNoteHeight(p),0)
    if(!Number.isSafeInteger(queueHeight))return fail('Queued footnote height exceeds safe integer geometry',document.document_id,'resource-limit')
   } else selected=candidate
   bodyHeight+=groupHeight
   if(onPage)prior.end=end
   else bodySlices.push({paragraph_id:paragraph.id,start:lineIndex,end})
   lineIndex=end
   if(lineIndex===shape.lines.length){paragraphIndex++;lineIndex=0}
  }
  if(!bodySlices.length){
   if(!pending.length)return fail('A body line group and the first footnote slice cannot fit together while honoring keep/widow constraints')
   selected=select(pending,height-activeRule.height)
   if(!selected.slices.length)return fail('A carried footnote line group cannot fit an empty page while honoring keep/widow constraints')
  }
  const noteHeight=selected.slices.length?selected.height+activeRule.height:0
  if(!Number.isSafeInteger(bodyHeight+noteHeight)||bodyHeight+noteHeight>height)return fail('Footnote flow geometry exceeds its page',document.document_id,'resource-limit')
  placedLines+=bodySlices.reduce((n,s)=>n+s.end-s.start,0)+selected.slices.reduce((n,s)=>n+s.end-s.start,0)+(selected.slices.length?1:0)
  if(placedLines>DOCX_NOTE_PAGINATION_LIMITS.maxNoteLines)return fail('Repeated footnote separators exceed the line budget',document.document_id,'resource-limit')
  pages.push({body:bodySlices,notes:selected.slices,note_height:noteHeight,...(selected.slices.length?{separator:activeRule.story}:{})})
  const ends=new Map(selected.slices.map(s=>[s.story_id,s.end]))
  pending=queue.flatMap(p=>{const start=ends.get(p.story_id)??p.start;return start<notes.get(p.story_id)!.lines.length?[{story_id:p.story_id,start}]:[]})
 }
 return {status:'planned',section_id:section.id,geometry:geometry.value,pages,notes}
}
