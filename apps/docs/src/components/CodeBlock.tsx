import { useState } from 'react'

const KEYWORDS = /\b(import|from|export|const|let|var|function|return|await|async|if|else|throw|new|type|interface|as|of|in|class|extends|package|func|err|nil|true|false|null|undefined)\b/g
const STRINGS = /('(?:\\'|[^'])*'|"(?:\\"|[^"])*"|`(?:\\`|[^`])*`)/g
const COMMENTS = /(\/\/.*$)/gm

function highlight(code: string): string {
  const escaped = code
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
  return escaped
    .replace(COMMENTS, '<span class="tok-cmt">$1</span>')
    .replace(STRINGS, '<span class="tok-str">$1</span>')
    .replace(KEYWORDS, '<span class="tok-kw">$1</span>')
}

export function CodeBlock({
  code,
  language = 'ts',
  title,
}: {
  code: string
  language?: 'ts' | 'bash' | 'go' | 'json' | 'tsx'
  title?: string
}) {
  const [copied, setCopied] = useState(false)
  const text = code.replace(/^\n/, '').replace(/\n$/, '')
  return (
    <div className="code">
      <header>
        <span>{title ?? language}</span>
        <button
          type="button"
          onClick={async () => {
            await navigator.clipboard.writeText(text)
            setCopied(true)
            window.setTimeout(() => setCopied(false), 1400)
          }}
        >
          {copied ? 'Copied' : 'Copy'}
        </button>
      </header>
      <pre>
        <code dangerouslySetInnerHTML={{ __html: highlight(text) }} />
      </pre>
    </div>
  )
}
