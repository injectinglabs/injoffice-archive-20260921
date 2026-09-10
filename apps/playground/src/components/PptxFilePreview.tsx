import { useEffect, useState, type ReactNode } from 'react'
import type { NativeElement, NativePptxDeck } from '@injoffice/pptx-native'
import { presetPath, type RenderPathCommand } from '@injoffice/pptx-render'
import { DsButton, DsField, DsSelect } from '../design-system/primitives'
import { previewColor, previewImage, previewIssue, previewSlideSize } from '../pptxPreview'

function geometry(commands: readonly RenderPathCommand[]): ReactNode {
  const first = commands[0]
  if (first?.kind === 'rect') return <rect width={first.rect.cx} height={first.rect.cy} />
  if (first?.kind === 'roundRect') return <rect width={first.rect.cx} height={first.rect.cy} rx={first.radiusEmu} />
  if (first?.kind === 'ellipse') return <ellipse cx={first.rect.cx / 2} cy={first.rect.cy / 2} rx={first.rect.cx / 2} ry={first.rect.cy / 2} />
  return <path d={commands.map((part) => part.kind === 'moveTo' ? `M${part.x} ${part.y}` : part.kind === 'lineTo' ? `L${part.x} ${part.y}` : part.kind === 'close' ? 'Z' : '').join(' ')} />
}

export default function PptxFilePreview({ deck }: { deck: NativePptxDeck }) {
  const [at, setAt] = useState(0)
  useEffect(() => setAt(0), [deck.documentId])
  const index = Math.min(at, Math.max(0, deck.slides.length - 1))
  const slide = deck.slides[index]
  const size = previewSlideSize(deck)
  if (!slide || !size) return <p className="ds-status">No previewable slide dimensions are available. The source file is unchanged.</p>
  const visible = slide.elements.slice(0, 500)
  const issues = visible.flatMap((element) => {
    const issue = previewIssue(element, deck.assets)
    return issue ? [`${element.name || element.kind}: ${issue}`] : []
  })
  if (slide.elements.length > visible.length) issues.push('Preview limited to 500 objects on this slide; remaining objects are preserved.')

  function object(element: NativeElement) {
    if (!size) return null
    const { x, y, cx, cy } = element.transform
    if (![x, y, cx, cy].every(Number.isFinite) || cx <= 0 || cy <= 0) return null
    const issue = previewIssue(element, deck.assets)
    const raster = element.kind === 'picture' ? previewImage(deck.assets.find((asset) => asset.id === element.assetId))
      : element.kind === 'chart' ? previewImage(deck.assets.find((asset) => asset.id === element.chart.previewAssetId)) : undefined
    const text = element.kind === 'text' || element.kind === 'shape' ? element.paragraphs : []
    const body = element.kind === 'text' || element.kind === 'shape' ? element.textBody : undefined
    const crop = element.kind === 'picture' ? element.crop : undefined
    return <g key={element.id} transform={`translate(${x * size.scale} ${y * size.scale})`}>
      <title>{element.name || element.kind}{issue ? `: ${issue}` : ''}</title>
      {issue ? <>
        <rect width={cx * size.scale} height={cy * size.scale} fill="#f5f6f6" stroke="#5a6560" strokeDasharray="4 3" />
        <text x="6" y="18" fontSize="12" fill="#141816">{element.kind} preview unavailable</text>
      </> : <>
        {element.kind === 'shape' && element.preset && <g transform={`scale(${size.scale})`} fill={previewColor(element.fill, 'none')} stroke={previewColor(element.stroke?.color, 'none')} strokeWidth={element.stroke?.widthEmu ?? 0}>{geometry(presetPath(element.preset, cx, cy))}</g>}
        {element.kind === 'connector' && <line x1={element.flipH ? cx * size.scale : 0} y1="0" x2={element.flipH ? 0 : cx * size.scale} y2={cy * size.scale} stroke={previewColor(element.stroke?.color, '#141816')} strokeWidth={Math.max(1, (element.stroke?.widthEmu ?? 9525) * size.scale)} />}
        {raster && (crop
          ? <svg width={cx * size.scale} height={cy * size.scale} viewBox={`${crop.left} ${crop.top} ${100_000 - crop.left - crop.right} ${100_000 - crop.top - crop.bottom}`} preserveAspectRatio="none" overflow="hidden"><image href={raster} width="100000" height="100000" preserveAspectRatio="none" /></svg>
          : <image href={raster} width={cx * size.scale} height={cy * size.scale} preserveAspectRatio="none" />)}
        {text.length > 0 && <foreignObject width={cx * size.scale} height={cy * size.scale} style={{ overflow: body ? 'visible' : 'hidden' }}>
          <div style={{ height: '100%', boxSizing: 'border-box', overflow: body ? 'visible' : 'hidden', color: '#141816', display: 'flex', flexDirection: 'column', justifyContent: body?.verticalAnchor === 'center' ? 'center' : body?.verticalAnchor === 'bottom' ? 'flex-end' : 'flex-start', padding: `${(body?.topInsetEmu ?? 0) * size.scale}px ${(body?.rightInsetEmu ?? 0) * size.scale}px ${(body?.bottomInsetEmu ?? 0) * size.scale}px ${(body?.leftInsetEmu ?? 0) * size.scale}px` }}>
            {text.map((paragraph, paragraphIndex) => <p key={paragraphIndex} style={{ margin: 0, paddingLeft: (paragraph.marginLeftEmu ?? 0) * size.scale, textIndent: (paragraph.indentEmu ?? 0) * size.scale, flexShrink: 0, lineHeight: 1.2, whiteSpace: body?.wrap === 'none' ? 'pre' : 'pre-wrap', overflowWrap: 'normal', textAlign: paragraph.align === 'center' ? 'center' : paragraph.align === 'right' ? 'right' : 'left', fontFamily: paragraph.runs[0]?.fontFamily ?? 'Arial, sans-serif', fontSize: (paragraph.runs[0]?.fontSizeHundredthPt ?? 1800) * 127 * size.scale }}>
              {paragraph.bullet && <span style={{ color: previewColor(paragraph.runs[0]?.color, '#141816'), fontWeight: paragraph.runs[0]?.bold ? 700 : 400, fontStyle: paragraph.runs[0]?.italic ? 'italic' : 'normal' }}>{`${paragraph.bulletCharacter ?? '•'} `}</span>}{paragraph.runs.map((run, runIndex) => <span key={runIndex} style={{ fontFamily: run.fontFamily ?? 'Arial, sans-serif', fontWeight: run.bold ? 700 : 400, fontStyle: run.italic ? 'italic' : 'normal', fontSize: (run.fontSizeHundredthPt ?? 1800) * 127 * size.scale, color: previewColor(run.color, '#141816') }}>{run.text}</span>)}
            </p>)}
          </div>
        </foreignObject>}
      </>}
    </g>
  }

  return <section aria-label="Presentation file preview" className="ds-panel">
    <div className="ds-workstrip">
      <DsButton variant="outlined" disabled={index === 0} onClick={() => setAt(index - 1)}>Previous slide</DsButton>
      <DsField label="Slide"><DsSelect value={index} onChange={(event) => setAt(Number(event.target.value))}>{deck.slides.map((item, i) => <option key={item.id} value={i}>{i + 1} of {deck.slides.length}</option>)}</DsSelect></DsField>
      <DsButton variant="outlined" disabled={index + 1 >= deck.slides.length} onClick={() => setAt(index + 1)}>Next slide</DsButton>
    </div>
    <p className="ds-status">Approximate file preview · browser fonts and text wrapping. Not PowerPoint-equivalent rendering. Editing remains limited to verified targets below.</p>
    <svg viewBox={`0 0 ${size.width} ${size.height}`} role="img" aria-label={`Approximate preview of slide ${index + 1}`} style={{ display: 'block', width: '100%', border: '1px solid var(--ds-line)', background: previewColor(slide.background) }}>
      {visible.map(object)}
    </svg>
    {issues.length > 0 && <details open><summary>{issues.length} preview limitation{issues.length === 1 ? '' : 's'}</summary><ul>{issues.map((issue, i) => <li key={i}>{issue}</li>)}</ul></details>}
    {slide.compatibility.diagnostics.length > 0 && <details><summary>Source compatibility warnings ({slide.compatibility.diagnostics.length})</summary><ul>{slide.compatibility.diagnostics.map((item, i) => <li key={i}>{item.message}</li>)}</ul></details>}
  </section>
}
