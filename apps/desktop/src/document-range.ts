export type DocumentTextRange={start_utf16:number;end_utf16:number;unsupported?:boolean;paragraph_id?:string}
// Offsets count authored UTF-16 text, excluding list markers and empty-run prompts.
export function paragraphTextOffset(lengths:number[],index:number,offset:number):number|undefined{
 if(!Number.isInteger(index)||index<0||index>=lengths.length||!Number.isInteger(offset)||offset<0||offset>lengths[index])return
 return lengths.slice(0,index).reduce((sum,length)=>sum+length,0)+offset
}
function spans(paragraph:HTMLElement){return Array.from(paragraph.querySelectorAll<HTMLElement>('[data-docx-run]'))}
function textLength(span:HTMLElement){return span.dataset.docxEmpty==='true'?0:(span.textContent??'').length}
export function captureParagraphSelection(paragraph:HTMLElement,range:Range):{start_utf16:number;end_utf16:number}|undefined{
 const runs=spans(paragraph),lengths=runs.map(textLength)
 const offset=(node:Node,position:number)=>{
  const index=runs.findIndex(run=>run===node||run.contains(node));if(index<0)return
  const prefix=range.cloneRange();prefix.selectNodeContents(runs[index]);prefix.setEnd(node,position)
  return paragraphTextOffset(lengths,index,prefix.toString().length)
 }
 const start_utf16=offset(range.startContainer,range.startOffset),end_utf16=offset(range.endContainer,range.endOffset)
 if(start_utf16===undefined||end_utf16===undefined||start_utf16>=end_utf16)return
 return {start_utf16,end_utf16}
}
export function restoreParagraphSelection(paragraph:HTMLElement,value:DocumentTextRange):void{
 const runs=spans(paragraph),document=paragraph.ownerDocument
 const point=(offset:number):{node:Node;offset:number}|undefined=>{
  for(const span of runs){const length=textLength(span);if(offset>length){offset-=length;continue}if(!length)continue
   const walker=document.createTreeWalker(span,4);let node:Node|null
   while((node=walker.nextNode())){const size=node.textContent?.length??0;if(offset<=size)return{node,offset};offset-=size}
  }
 }
 const start=point(value.start_utf16),end=point(value.end_utf16);if(!start||!end)return
 const selection=document.defaultView?.getSelection();if(!selection)return
 const current=selection.rangeCount?selection.getRangeAt(0):undefined
 if(current?.startContainer===start.node&&current.startOffset===start.offset&&current.endContainer===end.node&&current.endOffset===end.offset)return
 const range=document.createRange();range.setStart(start.node,start.offset);range.setEnd(end.node,end.offset);selection.removeAllRanges();selection.addRange(range)
}
