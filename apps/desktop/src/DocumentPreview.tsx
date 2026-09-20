import { useEffect, useId, useLayoutEffect, useRef, useState, type CSSProperties, type ReactNode } from 'react'
import {isRasterDataUrl} from './document-media'
import type { NativeDocxDocumentV1, NativeDocxParagraphV1, NativeDocxRunV1, NativeDocxStoryV1 } from '../../../packages/docs/src/nativeContract'
import { captureParagraphSelection, restoreParagraphSelection, type DocumentTextRange } from './document-range'
import { paragraphAppearance, paragraphListLabels, runAppearance } from './document-style'
import { editableDocxRuns } from '../../playground/src/docxRoundTrip'
import { nativeDocxParagraphText, nativeDocxRunText } from '../../playground/src/docsNativePreview'

interface DocumentPreviewProps {
  replaceImage?(id:string):void
  deleteImage(id:string):void
  images:Record<string,string>
  imageNotice:string
  document: NativeDocxDocumentV1
  selected: string
  draft: string
  textRange?: DocumentTextRange
  onTextRangeChange(value?: DocumentTextRange): void
  caretOffset?: number
  joinPrevious(): void
  insertLines(text: string, caret: number): void
  busy: boolean
  hasDraft: boolean
  zoom: number
  navigation: boolean
  choose(key: string,range?:DocumentTextRange): void
  updateDraft(value: string): void
  apply(): void
  cancel(): void
  onCompositionChange(value: boolean): void
}

function runIdentity(partName: string, runId: string) { return `${partName}\0${runId}` }
function baseParagraphId(paragraph: NativeDocxParagraphV1) { return `doc-paragraph-${encodeURIComponent(paragraph.anchor.part_name)}-${encodeURIComponent(paragraph.id)}` }
function headingLevel(paragraph: NativeDocxParagraphV1, document: NativeDocxDocumentV1): number | undefined {
  const level = paragraphAppearance(document, paragraph).paragraph.outline_level
  if (level !== undefined) return level < 9 ? level + 1 : undefined
  const match = /^heading[ _-]?([1-9])$/i.exec(paragraph.properties.paragraph_style_id ?? '')
  return match ? Number(match[1]) : undefined
}
function hexColor(value: string | undefined): string | undefined { return value && /^[0-9a-f]{6}$/i.test(value) ? `#${value}` : undefined }
function readableColor(color: string | undefined, background: string): string | undefined {
  const luminance = (hex: string) => {
    const channels = [1, 3, 5].map(offset => {
      const value = parseInt(hex.slice(offset, offset + 2), 16) / 255
      return value <= .04045 ? value / 12.92 : ((value + .055) / 1.055) ** 2.4
    })
    return channels[0]! * .2126 + channels[1]! * .7152 + channels[2]! * .0722
  }
  const contrast = (foreground: string) => {
    const left = luminance(foreground), right = luminance(background)
    return (Math.max(left, right) + .05) / (Math.min(left, right) + .05)
  }
  // This affects only the preview: an unavailable background must not hide authored text.
  if (contrast(color ?? '#20242b') >= 4.5) return color
  return contrast('#20242b') >= contrast('#ffffff') ? '#20242b' : '#ffffff'
}
function bodyParagraphs(story: NativeDocxStoryV1) {
  return story.blocks.flatMap(block => block.paragraph ? [block.paragraph] : block.table?.rows.flatMap(row => row.cells.flatMap(cell => cell.paragraphs)) ?? [])
}

export default function DocumentPreview(props: DocumentPreviewProps) {
  const { document, selected, choose, busy, hasDraft, zoom, navigation } = props
  const previewId=useId(),root=useRef<HTMLDivElement>(null)
  const paragraphId=(value:NativeDocxParagraphV1)=>`${previewId}-${baseParagraphId(value)}`
  const findParagraph=(id:string)=>Array.from(root.current?.querySelectorAll<HTMLElement>('[data-docx-paragraph]')??[]).find(node=>node.id===id)
  const targets = new Map(editableDocxRuns(document).map(target => [runIdentity(target.partName, target.runId), target]))
  const headings = bodyParagraphs(document.body).filter(paragraph => headingLevel(paragraph, document) !== undefined)
  useLayoutEffect(()=>{
    if(!props.textRange?.paragraph_id)return
    const value=bodyParagraphs(document.body).find(item=>item.id===props.textRange!.paragraph_id)
    const node=value&&findParagraph(paragraphId(value))
    if(node)restoreParagraphSelection(node,props.textRange)
  },[document,selected,props.textRange])
  const [currentHeading, setCurrentHeading] = useState('')
  const listLabels = new Map([document.body,...document.headers,...document.footers,...document.notes,...document.comment_stories].flatMap(story => [...paragraphListLabels(document,bodyParagraphs(story))]));
  function paragraph(value: NativeDocxParagraphV1, background = '#ffffff'): ReactNode {
    const level = headingLevel(value, document)
    const isTitle = /^title$/i.test(value.properties.paragraph_style_id ?? '')
    const properties = paragraphAppearance(document, value).paragraph
    const label = listLabels.get(runIdentity(value.anchor.part_name,value.id))
    const alignment = properties.alignment
    const style: CSSProperties = { textAlign: alignment === 'both' || alignment === 'distribute' ? 'justify' : alignment,
      marginTop: properties.spacing_before_twips === undefined ? undefined : `${properties.spacing_before_twips / 20}pt`, marginBottom: properties.spacing_after_twips === undefined ? undefined : `${properties.spacing_after_twips / 20}pt`,
      marginLeft: properties.indent_left_twips === undefined ? undefined : `${properties.indent_left_twips / 20}pt`, marginRight: properties.indent_right_twips === undefined ? undefined : `${properties.indent_right_twips / 20}pt`,
      textIndent: properties.hanging_twips !== undefined ? `${-properties.hanging_twips / 20}pt` : properties.first_line_twips === undefined ? undefined : `${properties.first_line_twips / 20}pt`,
      lineHeight: properties.line_spacing === undefined ? undefined : properties.line_rule && properties.line_rule !== 'auto' ? `${properties.line_spacing / 20}pt` : properties.line_spacing / 240,
    }
    return <p key={runIdentity(value.anchor.part_name, value.id)} id={paragraphId(value)} onMouseUp={event=>{
      const selection=window.getSelection();if(busy||!value.can_format_range||value.runs.some(run=>runAppearance(document,value,run).hidden)||!selection?.rangeCount)return
      const range=selection.getRangeAt(0);if(!event.currentTarget.contains(range.commonAncestorContainer))return
      const captured=captureParagraphSelection(event.currentTarget,range);if(!captured)return
      const first=value.runs.find(run=>targets.has(runIdentity(run.anchor.part_name,run.id)))
      if(!first)return
      const next={...captured,paragraph_id:value.id}
      if(value.runs.some(run=>targets.get(runIdentity(run.anchor.part_name,run.id))?.key===selected))props.onTextRangeChange(next)
      else choose(targets.get(runIdentity(first.anchor.part_name,first.id))!.key,next)
    }} data-docx-paragraph={value.id} data-range-editable={value.can_format_range&&!value.runs.some(run=>runAppearance(document,value,run).hidden)?true:undefined} tabIndex={-1} className={`office-paragraph ${level ? `office-heading office-heading-${level}` : ''} ${isTitle ? 'office-doc-title' : ''}`} style={style} role={level ? 'heading' : undefined} aria-level={level}>
      {label && <span className="office-list-marker" contentEditable={false} style={{fontFamily:label.properties?.font_family,fontSize:label.properties?.font_size_half_points ? `${label.properties.font_size_half_points / 2}pt` : undefined}}>{label.text}{label.suffix === 'nothing' ? '' : '\u00a0'}</span>}
      {value.runs.map(run => {
        const key = runIdentity(run.anchor.part_name, run.id)
        const target = targets.get(key)
        return <DocumentRun replaceImage={props.replaceImage} deleteImage={props.deleteImage} image={run.drawing?props.images[run.drawing.id]:undefined} key={key} run={{...run, properties:runAppearance(document, value, run)}} background={background} active={!!target && selected === target.key} editable={!!target} disabled={busy} draft={props.draft} textRange={props.textRange} onTextRangeChange={props.onTextRangeChange} caretOffset={props.caretOffset} joinPrevious={props.joinPrevious} insertLines={props.insertLines} activate={() => target && choose(target.key)} updateDraft={props.updateDraft} apply={props.apply} cancel={props.cancel} onCompositionChange={props.onCompositionChange} />
      })}
    </p>
  }
  function story(value: NativeDocxStoryV1) {
    return value.blocks.map((block, index) => block.paragraph ? paragraph(block.paragraph) : block.table ? <table className="office-doc-table" key={index} style={{width:block.table.width_twips === undefined ? undefined : `${block.table.width_twips / 20}pt`,tableLayout:block.table.layout}}>{block.table.grid_widths_twips && <colgroup>{block.table.grid_widths_twips.map((width,index)=><col key={index} style={{width:`${width / 20}pt`}} />)}</colgroup>}<tbody>{block.table.rows.map((row, r) => <tr key={r}>{row.cells.map((cell, c) => {
      const background = hexColor(cell.shading_rgb) ?? '#ffffff'
      return <td key={c} colSpan={cell.grid_span} style={{ backgroundColor: background }}>{cell.paragraphs.map(value => paragraph(value, background))}</td>
    })}</tr>)}</tbody></table> : <p key={index} className="office-preserved-content">Content preserved in the original file</p>)
  }
  const section=document.sections.length===1?document.sections[0]:undefined
  const page=section?.edit_policy?.allowed_operations.includes('section.page.patch')?section.page:undefined
  const paperStyle:CSSProperties|undefined=page?{boxSizing:'border-box',width:`${page.width_twips/20}pt`,minHeight:`${page.height_twips/20}pt`,paddingTop:`${page.margins.top_twips/20}pt`,paddingRight:`${page.margins.right_twips/20}pt`,paddingBottom:`${page.margins.bottom_twips/20}pt`,paddingLeft:`${(page.margins.left_twips+page.margins.gutter_twips)/20}pt`}:undefined
  const secondaryStories = [
    ...document.headers.map((value, index) => ({ value, label: `Header ${index + 1}` })),
    ...document.footers.map((value, index) => ({ value, label: `Footer ${index + 1}` })),
    ...document.notes.map((value, index) => ({ value, label: `Note ${index + 1}` })),
    ...document.comment_stories.map((value, index) => ({ value, label: `Comment ${index + 1}` })),
  ]
  return <div ref={root} className="office-document-layout">
    {navigation && <nav className="office-navigation" aria-label="Document headings">
      <h2>Navigation</h2><div className="office-navigation-tab">Headings</div>
      {headings.length ? <ol>{headings.map(heading => <li key={paragraphId(heading)} style={{ paddingLeft: `${Math.min(4, (headingLevel(heading, document) ?? 1) - 1) * 12}px` }}><button aria-current={currentHeading === paragraphId(heading) ? 'location' : undefined} onClick={() => {
        const id = paragraphId(heading)
        setCurrentHeading(id)
        const element = findParagraph(id)
        element?.scrollIntoView({ block: 'start', behavior: 'instant' })
        element?.focus({ preventScroll: true })
      }}>{nativeDocxParagraphText(heading) || 'Untitled heading'}</button></li>)}</ol> : <p>Headings in your document appear here.</p>}
    </nav>}
    <div className="office-document-canvas">
      <p className="office-document-note">Flowing document preview. Original page layout is preserved in the file.</p>
      {props.imageNotice&&<p className="office-document-note">{props.imageNotice}</p>}
      <div className="office-document-zoom" style={{ zoom }}>
        <article style={paperStyle} className="office-paper" aria-label="Document content">{story(document.body)}</article>
        {secondaryStories.map(({ value, label }) => <section className="office-story" key={label} aria-label={label}><h2>{label}</h2><div className="office-story-content">{story(value)}</div></section>)}
      </div>
    </div>
  </div>
}

interface DocumentRunProps {
  replaceImage?(id:string):void
  deleteImage(id:string):void
  image?:string
  run: NativeDocxRunV1
  background: string
  active: boolean
  editable: boolean
  disabled: boolean
  draft: string
  textRange?: DocumentTextRange
  onTextRangeChange(value?: DocumentTextRange): void
  caretOffset?: number
  joinPrevious(): void
  insertLines(text: string, caret: number): void
  activate(): void
  updateDraft(value: string): void
  apply(): void
  cancel(): void
  onCompositionChange(value: boolean): void
}

function DocumentRun({ replaceImage, deleteImage, image, run, background, active, editable, disabled, draft, textRange, onTextRangeChange, caretOffset, insertLines, joinPrevious, activate, updateDraft, apply, cancel, onCompositionChange }: DocumentRunProps) {
  const element = useRef<HTMLSpanElement>(null)
  const point = useRef<{ x: number; y: number } | undefined>(undefined)
  const composing = useRef(false)
  const [hint, setHint] = useState('')
  const initialText = run.text ?? ''
  // React never owns children of the editable span. Ordinary rerenders must not reset its DOM/caret.
  useLayoutEffect(() => {
    if (!active || !element.current) return
    const node = element.current
    node.textContent = initialText
    node.focus({ preventScroll: true })
    const selection = window.getSelection()
    const range = window.document.createRange()
    const clicked = point.current ? window.document.caretRangeFromPoint(point.current.x, point.current.y) : null
    if (clicked && node.contains(clicked.startContainer)) {
      range.setStart(clicked.startContainer, clicked.startOffset)
      range.collapse(true)
    } else if (textRange && !textRange.unsupported && !textRange.paragraph_id && node.firstChild) { range.setStart(node.firstChild,Math.min(textRange.start_utf16,node.textContent?.length??0));range.setEnd(node.firstChild,Math.min(textRange.end_utf16,node.textContent?.length??0)) } else if (caretOffset !== undefined && node.firstChild) { range.setStart(node.firstChild, Math.min(caretOffset, node.textContent?.length ?? 0)); range.collapse(true) } else { range.selectNodeContents(node); range.collapse(false) }
    selection?.removeAllRanges(); selection?.addRange(range)
    point.current = undefined
    setHint('')
  }, [active, initialText])
  // Cancel / accepted values may update while active, but normal input already equals draft.
  useLayoutEffect(() => {
    if (active && element.current && !composing.current && element.current.textContent !== draft) element.current.textContent = draft
  }, [active, draft])
  const selectionCallback=useRef(onTextRangeChange);selectionCallback.current=onTextRangeChange
  useEffect(()=>{
    if(!active)return
    const change=()=>{
      const node=element.current,selection=window.getSelection()
      if(!node||!selection?.rangeCount)return
      const range=selection.getRangeAt(0)
      // Toolbar focus is outside the editable span; keep the captured range.
      if(!node.contains(range.commonAncestorContainer)){
        if(!range.collapsed&&range.intersectsNode(node)){
          const paragraph=node.closest<HTMLElement>('[data-docx-paragraph]')
          const captured=paragraph?.dataset.rangeEditable==='true'&&paragraph.contains(range.commonAncestorContainer)?captureParagraphSelection(paragraph,range):undefined
          selectionCallback.current(captured?{...captured,paragraph_id:paragraph!.dataset.docxParagraph}:{start_utf16:0,end_utf16:range.toString().length,unsupported:true})
        }
        return
      }
      if(range.collapsed){selectionCallback.current(undefined);return}
      const prefix=range.cloneRange();prefix.selectNodeContents(node);prefix.setEnd(range.startContainer,range.startOffset)
      const start_utf16=prefix.toString().length
      selectionCallback.current({start_utf16,end_utf16:start_utf16+range.toString().length})
    }
    window.document.addEventListener('selectionchange',change)
    return()=>window.document.removeEventListener('selectionchange',change)
  },[active])
  function insertParagraphText(text: string) {
    const selection = window.getSelection(), node = element.current
    if (!node || !selection?.rangeCount) return
    const range = selection.getRangeAt(0)
    if (!node.contains(range.commonAncestorContainer)) return
    const before = range.cloneRange(); before.selectNodeContents(node); before.setEnd(range.startContainer, range.startOffset)
    const after = range.cloneRange(); after.selectNodeContents(node); after.setStart(range.endContainer, range.endOffset)
    insertLines(before.toString() + text + after.toString(), text.split(/\r\n?|\n/).at(-1)!.length)
  }
  const properties = run.properties
  const style: CSSProperties = {
    fontWeight: properties?.bold ? 'bold' : properties?.bold === false ? 'normal' : undefined,
    fontStyle: properties?.italic ? 'italic' : undefined,
    textDecoration: properties?.underline && properties.underline !== 'none' ? 'underline' : undefined,
    fontFamily: properties?.font_family,
    fontSize: properties?.font_size_half_points ? `${properties.font_size_half_points / 2}pt` : undefined,
    color: readableColor(hexColor(properties?.color), background),
  }
  const text = nativeDocxRunText(run)
  if (run.properties?.hidden) return null
  if(run.drawing)return <span className="office-inline-image" contentEditable={false}>{isRasterDataUrl(image)?<img src={image} alt={run.drawing.alt_text??run.drawing.name??'Embedded image'} style={{width:`${run.drawing.width_emu/12700}pt`,maxWidth:'100%',height:'auto',aspectRatio:`${run.drawing.width_emu}/${run.drawing.height_emu}`,verticalAlign:'middle'}} draggable={false}/>:<span className="office-preserved-image">Image preserved</span>}<span className="office-image-actions">{replaceImage&&run.drawing.edit_policy.allowed_operations.includes('drawing.replace')&&<button className="office-image-replace" disabled={disabled} aria-label={`Replace picture ${run.drawing.alt_text??run.drawing.name??''}`} onClick={()=>replaceImage(run.drawing!.id)}>Replace picture…</button>}{run.drawing.edit_policy.allowed_operations.includes('block.delete')&&<button className="office-image-delete" disabled={disabled} aria-label={`Delete image ${run.drawing.alt_text??run.drawing.name??''}`} onClick={()=>deleteImage(run.drawing!.id)}>Delete image</button>}</span></span>
  if (!editable) return <span style={style}>{text || (run.kind === 'drawing' ? <span className="office-preserved-image">Image</span> : '')}</span>
  if (!active) return <span role="button" tabIndex={disabled ? -1 : 0} aria-disabled={disabled} className="office-text-run" data-docx-run={run.id} data-docx-empty={!text?true:undefined} data-hyperlink={run.hyperlink?true:undefined} style={style} title={run.hyperlink?`Link: ${run.hyperlink.url} · Click to edit text`:"Click to edit text"} onClick={event => {
    if (disabled) return
    const selection=window.getSelection()
    if(selection?.rangeCount&&!selection.isCollapsed&&selection.getRangeAt(0).intersectsNode(event.currentTarget))return
    point.current = { x: event.clientX, y: event.clientY }; activate()
  }} onKeyDown={event => { if (!disabled && (event.key === 'Enter' || event.key === ' ')) { event.preventDefault(); point.current = undefined; activate() } }}>{text || <span className="office-empty-run">Start typing…</span>}</span>
  return <><span ref={element} className="office-inline-input" data-docx-run={run.id} data-hyperlink={run.hyperlink?true:undefined} title={run.hyperlink?.url} contentEditable={disabled ? false : 'plaintext-only'} suppressContentEditableWarning role="textbox" aria-label="Edit document text" aria-multiline="false" data-placeholder="Start typing…" style={style} spellCheck onInput={event => updateDraft(event.currentTarget.textContent ?? '')}
    onCompositionStart={() => { composing.current = true; onCompositionChange(true) }}
    onCompositionEnd={event => { composing.current = false; updateDraft(event.currentTarget.textContent ?? ''); onCompositionChange(false) }}
    onCut={event=>{const selection=window.getSelection();if(selection?.rangeCount&&!element.current?.contains(selection.getRangeAt(0).commonAncestorContainer)){event.preventDefault();setHint('Select text inside one segment to cut.')}}}
    onBeforeInput={event => { const selection=window.getSelection();if(selection?.rangeCount&&!element.current?.contains(selection.getRangeAt(0).commonAncestorContainer)){event.preventDefault();setHint('Select text inside one segment to type. Formatting can span segments.');return} const input = event.nativeEvent as InputEvent; if (input.inputType === 'insertParagraph' || input.inputType === 'insertLineBreak') { event.preventDefault(); setHint('Paragraph breaks are not supported in this text segment.') } }}
    onPaste={event => {
      event.preventDefault()
      const text = event.clipboardData.getData('text/plain')
      if (/[\r\n]/.test(text)) { insertParagraphText(text); return }
      if (/[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(text)) { setHint('This text contains unsupported control characters.'); return }
      // Insert a text node into the active selection; clipboard HTML is never interpreted.
      const selection = window.getSelection()
      if (!selection?.rangeCount || !element.current) return
      const range = selection.getRangeAt(0)
      if (!element.current.contains(range.commonAncestorContainer)) return
      range.deleteContents()
      const node = window.document.createTextNode(text)
      range.insertNode(node); range.setStartAfter(node); range.collapse(true)
      selection.removeAllRanges(); selection.addRange(range)
      updateDraft(element.current.textContent ?? ''); setHint('')
    }}
    onDrop={event => { event.preventDefault(); setHint('Use paste to insert plain text.') }}
    onKeyDown={event => {
      if (composing.current || event.nativeEvent.isComposing) return
      if (event.key === 'Backspace') {
        const selection = window.getSelection(), node = element.current
        if (node && selection?.isCollapsed && selection.rangeCount) {
          const range = selection.getRangeAt(0)
          if (node.contains(range.commonAncestorContainer)) { const before = range.cloneRange(); before.selectNodeContents(node); before.setEnd(range.startContainer, range.startOffset); if (!before.toString()) { event.preventDefault(); joinPrevious(); return } }
        }
      }
      if (event.key === 'Escape') { event.preventDefault(); cancel() }
      else if (event.key === 'Enter') { event.preventDefault(); if (event.ctrlKey || event.metaKey) apply(); else insertParagraphText('\n') }
    }} />{hint && <span className="office-inline-hint" role="status">{hint}</span>}</>
}
