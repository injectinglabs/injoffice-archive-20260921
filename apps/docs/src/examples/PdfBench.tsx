import { useEffect, useState } from 'react'
import { PDFDocument, StandardFonts } from 'pdf-lib'
import { applyPageOps, readInfo, type PdfDocumentInfo } from '@injoffice/pdf/browser'
import { LiveBench } from '../components/LiveBench'

async function samplePdf(): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const font = await doc.embedFont(StandardFonts.Helvetica)
  const cover = doc.addPage([400, 520])
  cover.drawText('InjOffice', { x: 48, y: 440, size: 28, font })
  cover.drawText('Page 1', { x: 48, y: 400, size: 14, font })
  const second = doc.addPage([400, 520])
  second.drawText('Page 2', { x: 48, y: 440, size: 28, font })
  return new Uint8Array(await doc.save())
}

function describe(info: PdfDocumentInfo): string {
  return info.pages.map((page) => `${page.index}:${page.rotation}° ${Math.round(page.width)}×${Math.round(page.height)}`).join(' · ')
}

export function PdfBench() {
  const [info, setInfo] = useState<PdfDocumentInfo | null>(null)
  const [error, setError] = useState('')
  const [busy, setBusy] = useState(false)
  const [bytes, setBytes] = useState<Uint8Array | null>(null)

  useEffect(() => {
    let cancelled = false
    void samplePdf().then(async (source) => {
      const next = await readInfo(source)
      if (cancelled) return
      setBytes(source)
      setInfo(next)
    }).catch((reason) => {
      if (!cancelled) setError(reason instanceof Error ? reason.message : String(reason))
    })
    return () => { cancelled = true }
  }, [])

  const rotate = async () => {
    if (!bytes || !info) return
    setBusy(true)
    setError('')
    try {
      const rotated = await applyPageOps(bytes, [{ type: 'rotate', pages: [1], degrees: 90 }])
      setBytes(rotated)
      setInfo(await readInfo(rotated))
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
    } finally {
      setBusy(false)
    }
  }

  return (
    <LiveBench title="Live example" hint="@injoffice/pdf/browser · readInfo + applyPageOps">
      <div className="bench-controls">
        <button type="button" className="bench-button" onClick={() => void rotate()} disabled={busy || !bytes}>Rotate page 1 by 90°</button>
      </div>
      {error ? <p><span className="badge badge--patch">refused</span> {error}</p> : null}
      {info ? <p><span className="badge badge--pass">{info.pageCount} pages</span> {describe(info)}</p> : <p>Creating a two-page sample…</p>}
    </LiveBench>
  )
}
