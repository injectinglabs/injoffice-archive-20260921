// Header/footer token evaluator (D6), reusing the SAME &-escape-code
// convention Sheets print already writes into oddHeader/oddFooter XML
// (go/xlsxpatch/printwrite.go: PrintSetup.HeaderCenter/FooterCenter are raw
// strings containing literal &P/&N/&D/&F, which Excel/LibreOffice evaluate at
// render time — there is no shared parser function to import, it's a string
// convention). Docs pagination renders client-side in the browser (unlike
// xlsxpatch, which defers evaluation to Excel/LibreOffice), so this module is
// the evaluator: same token characters, computed here instead.
//
// Supported tokens (case-insensitive, matching Excel's header/footer codes):
//   &P  current page number
//   &N  total page count
//   &D  date (short, e.g. 8/25/2026)
//   &T  time (e.g. 3:04 PM)
//   &F  file name
// Any other "&X" is passed through literally (unknown codes are not silently
// eaten — matches Excel's own behavior of leaving unrecognized codes alone).

import { asciiUpperNative } from './nativeDeterminism.js'

export interface HeaderFooterContext {
  page: number
  totalPages: number
  date?: Date
  filename?: string
}

const TOKEN_RE = /&([PNDTF])/gi

function formatDate(d: Date): string {
  return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`
}

function formatTime(d: Date): string {
  let h = d.getHours()
  const m = d.getMinutes()
  const ampm = h >= 12 ? 'PM' : 'AM'
  h = h % 12 || 12
  return `${h}:${String(m).padStart(2, '0')} ${ampm}`
}

/** Substitute &P/&N/&D/&T/&F tokens in a header/footer template string. */
export function renderHeaderFooterTemplate(template: string, ctx: HeaderFooterContext): string {
  if (!template) return ''
  const date = ctx.date ?? new Date()
  return template.replace(TOKEN_RE, (whole, code: string) => {
    switch (asciiUpperNative(code)) {
      case 'P':
        return String(ctx.page)
      case 'N':
        return String(ctx.totalPages)
      case 'D':
        return formatDate(date)
      case 'T':
        return formatTime(date)
      case 'F':
        return ctx.filename ?? ''
      default:
        return whole
    }
  })
}

/** True if a template string contains at least one recognized token. */
export function hasHeaderFooterTokens(template: string): boolean {
  TOKEN_RE.lastIndex = 0
  return TOKEN_RE.test(template)
}
