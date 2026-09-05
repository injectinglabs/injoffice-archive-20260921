import { lazy, Suspense, useEffect, useState } from 'react'
import {
  parseCollabQuery,
  writeCollabQuery,
  type CollabFormat,
} from './collabScope'
import { SIM_FORMAT_LABEL, SimFormatTabs } from './collabSimChrome'

const CollabSheetsPanel = lazy(() => import('./collab/sheets'))
const CollabSlidesPanel = lazy(() => import('./collab/slides'))
const CollabDocsPanel = lazy(() => import('./collab/docs'))
const CollabPdfPanel = lazy(() => import('./collab/pdf'))

const FORMAT_HINT: Record<CollabFormat, string> = {
  sheets: 'Univer sheet protocol proof',
  slides: 'DeckSpec protocol proof',
  docs: 'ProseMirror document protocol proof',
  pdf: 'PDF annotation protocol proof',
}

export function CollabDemo() {
  const [format, setFormat] = useState<CollabFormat>(() => parseCollabQuery(window.location.href).format)

  useEffect(() => {
    const { artifact } = parseCollabQuery(window.location.href)
    writeCollabQuery({ artifact, format })
  }, [format])

  const selectFormat = (next: CollabFormat) => {
    const { artifact } = parseCollabQuery(window.location.href)
    writeCollabQuery({ artifact, format: next })
    setFormat(next)
  }

  return (
    <div className="collab-demo workbench-surface">
      <SimFormatTabs format={format} onFormat={selectFormat} hint={FORMAT_HINT[format]} />
      <Suspense fallback={<div className="demo-loading" role="status">Opening {SIM_FORMAT_LABEL[format]} collaboration…</div>}>
        {format === 'sheets' ? <CollabSheetsPanel /> : null}
        {format === 'slides' ? <CollabSlidesPanel /> : null}
        {format === 'docs' ? <CollabDocsPanel /> : null}
        {format === 'pdf' ? <CollabPdfPanel /> : null}
      </Suspense>
    </div>
  )
}
