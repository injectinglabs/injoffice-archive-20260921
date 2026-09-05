import { useEffect, useRef } from 'react'
import { LocaleType } from '@univerjs/core'
import { createOssUniver, createSheetsPresetBundle } from '@injoffice/univer-sheets/browser'
import '@injoffice/univer-sheets/styles.css'
import { LiveBench } from '../components/LiveBench'

export function EditorBench() {
  const hostRef = useRef<HTMLDivElement | null>(null)
  useEffect(() => {
    const host = hostRef.current
    if (!host) return
    const bundle = createSheetsPresetBundle({
      container: host,
      features: { threadComments: false, notes: false },
    })
    const { univer, univerAPI } = createOssUniver({
      locale: LocaleType.EN_US,
      locales: { [LocaleType.EN_US]: bundle.locale },
      presets: bundle.presets,
    })
    univerAPI.createWorkbook({
      id: 'docs-demo',
      name: 'docs-demo.xlsx',
      sheets: {
        s1: {
          id: 's1',
          name: 'Plan',
          cellData: {
            0: { 0: { v: 'Item' }, 1: { v: 'Owner' }, 2: { v: 'Status' } },
            1: { 0: { v: 'Native XLSX' }, 1: { v: 'Mira' }, 2: { v: 'Live' } },
            2: { 0: { v: 'Charts' }, 1: { v: 'Noah' }, 2: { v: 'Ready' } },
            3: { 0: { f: '=COUNTA(A2:A3)' }, 1: { v: '' }, 2: { v: '' } },
          },
          columnData: { 0: { w: 140 }, 1: { w: 100 }, 2: { w: 100 } },
        },
      },
    })
    return () => {
      ;(univer as { dispose?: () => void }).dispose?.()
    }
  }, [])
  return (
    <LiveBench title="Live example" hint="@injoffice/univer-sheets · createOssUniver">
      <div ref={hostRef} className="editor-host" id="univer-root" />
    </LiveBench>
  )
}
