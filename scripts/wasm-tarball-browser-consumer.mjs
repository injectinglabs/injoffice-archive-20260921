const marker = `tarball-browser-${Date.now().toString(36)}`

report({ state: 'running', current: 'xlsx' })

try {
  const results = []
  results.push(await exerciseXlsx(marker))
  report({ state: 'running', current: 'pptx', results })
  results.push(await exercisePptx(marker))
  report({ state: 'running', current: 'docx', results })
  results.push(await exerciseDocx(marker))
  report({ state: 'passed', results })
} catch (error) {
  report({
    state: 'failed',
    error: error instanceof Error ? `${error.name}: ${error.message}\n${error.stack ?? ''}` : String(error),
  })
}

async function exerciseXlsx(value) {
  const { createXlsxWasmClient } = await import('@injoffice/xlsx-wasm')
  const original = await fixture('pass-excel-defaults.xlsx')
  const client = createXlsxWasmClient()
  try {
    const workbook = await client.extract(original)
    const sheet = workbook.sheets[0]
    if (!sheet) throw new Error('XLSX fixture has no sheet')
    const target = sheet.cells.find((cell) => cell.editable)
    if (!target) throw new Error('XLSX fixture has no editable cell')
    const produced = await client.apply(original, workbook, {
      expected_revision: workbook.revision,
      cells: [{
        operation_id: 'tarball-browser-cell',
        sheet_id: sheet.id,
        kind: 'cell.set_value',
        cell: { row: target.row, column: target.column },
        value,
      }],
    })
    const reopened = await client.extract(produced)
    const cell = reopened.sheets[0]?.cells.find((candidate) =>
      candidate.row === target.row && candidate.column === target.column,
    )
    if (cell?.value?.kind !== 'string' || cell.value.text !== value) {
      throw new Error(`XLSX readback mismatch: ${JSON.stringify(cell?.value)}`)
    }
    assertRevisionChanged('XLSX', workbook.source.package_sha256, reopened.source.package_sha256)
    return { format: 'xlsx', value, bytes: produced.byteLength }
  } finally {
    client.terminate()
  }
}

async function exerciseDocx(value) {
  const { createDocxWasmClient } = await import('@injoffice/docx-wasm')
  const original = await fixture('docx-transitional-common.docx')
  const client = createDocxWasmClient()
  try {
    const document = await client.extract(original)
    const run = firstDocxTextRun(document)
    const produced = await client.apply(original, document, {
      protocol: 'injoffice.office.mutations',
      version: 1,
      format: 'docx',
      mutation_id: 'tarball-browser-run',
      expected_revision: document.source.package_sha256,
      payload: {
        mutations: [{
          target_kind: 'run',
          target_id: run.id,
          expected_xml_sha256: run.anchor.xml_sha256,
          text: value,
        }],
      },
    })
    const reopened = await client.extract(produced)
    const reopenedRun = docxTextRuns(reopened).find((candidate) => candidate.id === run.id)
    if (!reopenedRun) throw new Error(`DOCX readback lost targeted run ${run.id}`)
    if (reopenedRun.anchor.part_name !== run.anchor.part_name || reopenedRun.anchor.path !== run.anchor.path) {
      throw new Error(`DOCX targeted run ${run.id} moved to a different source anchor`)
    }
    if (reopenedRun.text !== value) throw new Error(`DOCX targeted run readback mismatch: ${JSON.stringify(reopenedRun.text)}`)
    if (reopenedRun.anchor.xml_sha256 === run.anchor.xml_sha256) throw new Error('DOCX targeted run source fingerprint did not change')
    assertRevisionChanged('DOCX', document.source.package_sha256, reopened.source.package_sha256)
    return { format: 'docx', value, bytes: produced.byteLength }
  } finally {
    client.terminate()
  }
}

async function exercisePptx(value) {
  const { createPptxWasmClient } = await import('@injoffice/pptx-wasm')
  const original = await fixture('pptx-transitional-common.pptx')
  const client = createPptxWasmClient()
  try {
    const deck = await client.extract(original)
    const element = firstPptxTextElement(deck)
    const paragraphs = element.paragraphs.map((paragraph, paragraphIndex) => ({
      align: paragraph.align ?? 'left',
      level: paragraph.level ?? 0,
      bullet: false,
      runs: paragraph.runs.map((run, runIndex) => ({
        text: paragraphIndex === 0 && runIndex === 0 ? value : run.text,
        bold: run.bold ?? false,
        italic: run.italic ?? false,
        fontSizeHundredthPt: run.fontSizeHundredthPt ?? 1800,
        color: run.color ?? '000000',
        fontFamily: run.fontFamily ?? 'Arial',
      })),
    }))
    const produced = await client.apply(original, deck, {
      expectedSourceRevision: deck.sourceRevision,
      operations: [{
        operationId: 'tarball-browser-text',
        kind: 'text.replace',
        elementId: element.id,
        expectedFingerprintSha256: element.source.fingerprintSha256,
        paragraphs,
      }],
    })
    const reopened = await client.extract(produced)
    const reopenedElement = allPptxElements(reopened).find((candidate) => candidate.id === element.id)
    if (!reopenedElement || (reopenedElement.kind !== 'text' && reopenedElement.kind !== 'shape')) {
      throw new Error(`PPTX readback lost targeted element ${element.id}`)
    }
    if (!reopenedElement.source
      || reopenedElement.source.partName !== element.source.partName
      || reopenedElement.source.objectId !== element.source.objectId) {
      throw new Error(`PPTX targeted element ${element.id} moved to a different source anchor`)
    }
    if (reopenedElement.paragraphs?.[0]?.runs[0]?.text !== value) {
      throw new Error(`PPTX targeted element readback mismatch: ${JSON.stringify(reopenedElement.paragraphs?.[0]?.runs[0]?.text)}`)
    }
    if (reopenedElement.source.fingerprintSha256 === element.source.fingerprintSha256) {
      throw new Error('PPTX targeted element source fingerprint did not change')
    }
    assertRevisionChanged('PPTX', deck.sourceRevision, reopened.sourceRevision)
    return { format: 'pptx', value, bytes: produced.byteLength }
  } finally {
    client.terminate()
  }
}

async function fixture(name) {
  const response = await fetch(`${import.meta.env.BASE_URL}fixtures/${name}`)
  if (!response.ok) throw new Error(`fixture ${name} returned HTTP ${response.status}`)
  return new Uint8Array(await response.arrayBuffer())
}

function firstDocxTextRun(document) {
  const run = docxTextRuns(document)[0]
  if (!run) throw new Error('DOCX fixture has no editable text run')
  return run
}

function docxTextRuns(document) {
  const result = []
  const stories = [document.body, ...document.headers, ...document.footers, ...document.notes, ...document.comment_stories]
  for (const story of stories) {
    for (const block of story.blocks) {
      if (block.paragraph) result.push(...block.paragraph.runs.filter((run) => run.kind === 'text'))
      if (block.table) {
        for (const row of block.table.rows) {
          for (const cell of row.cells) {
            for (const paragraph of cell.paragraphs) result.push(...paragraph.runs.filter((run) => run.kind === 'text'))
          }
        }
      }
    }
  }
  return result
}

function firstPptxTextElement(deck) {
  const element = allPptxElements(deck).find((candidate) =>
    (candidate.kind === 'text' || candidate.kind === 'shape')
      && candidate.source
      && candidate.paragraphs?.some((paragraph) => paragraph.runs.length > 0),
  )
  if (!element) throw new Error('PPTX fixture has no editable text element')
  return element
}

function allPptxElements(deck) {
  const result = []
  const visit = (elements) => {
    for (const element of elements) {
      result.push(element)
      if (element.kind === 'group') visit(element.children)
    }
  }
  for (const slide of deck.slides) visit(slide.elements)
  return result
}

function assertRevisionChanged(format, before, after) {
  if (!before || !after || before === after) throw new Error(`${format} output revision did not change`)
}

function report(value) {
  globalThis.__injofficeWasmTarballSmoke = value
  document.body.textContent = JSON.stringify(value)
}
