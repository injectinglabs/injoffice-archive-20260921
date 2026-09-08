import { lazy, Suspense, useEffect, useState } from 'react'
import {
  parseCollabQuery,
  writeCollabQuery,
  type CollabFormat,
} from './collabScope'
import { SIM_FORMAT_LABEL, SimFormatTabs } from './collabSimChrome'
import { CollabCompositionContext, type CollabCompositionProps } from './collabComposition'

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

export function CollabDemo({ fixedFormat, initialHash }: CollabCompositionProps = {}) {
  const [selectedFormat, setFormat] = useState<CollabFormat>(() => fixedFormat ?? parseCollabQuery(initialHash ?? window.location.href).format)
  const format = fixedFormat ?? selectedFormat

  useEffect(() => {
    if (fixedFormat !== undefined) return
    const { artifact } = parseCollabQuery(window.location.href)
    writeCollabQuery({ artifact, format })
  }, [format, fixedFormat])

  const selectFormat = (next: CollabFormat) => {
    if (fixedFormat !== undefined) return
    const { artifact } = parseCollabQuery(window.location.href)
    writeCollabQuery({ artifact, format: next })
    setFormat(next)
  }

  return (
    <CollabCompositionContext.Provider value={{ fixedFormat, initialHash }}>
    <div className="collab-demo workbench-surface ds">
      {fixedFormat === undefined && <SimFormatTabs format={format} onFormat={selectFormat} hint={FORMAT_HINT[format]} />}
      <Suspense fallback={<div className="demo-loading" role="status">Opening {SIM_FORMAT_LABEL[format]} collaboration…</div>}>
        {format === 'sheets' ? <CollabSheetsPanel /> : null}
        {format === 'slides' ? <CollabSlidesPanel /> : null}
        {format === 'docs' ? <CollabDocsPanel /> : null}
        {format === 'pdf' ? <CollabPdfPanel /> : null}
      </Suspense>
    </div>
    </CollabCompositionContext.Provider>
  )
}
