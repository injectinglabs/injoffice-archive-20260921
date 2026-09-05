// The formula-coverage audit (Phase 7): every target function evaluated in
// the REAL Univer engine, headless. Not a list comparison — declared enums
// prove nothing about computation; #NAME? and thrown parses do. The test
// FAILS when coverage regresses below the audited baseline, so engine
// upgrades can't silently drop functions agents rely on.
import { describe, expect, it } from 'vitest'
import { LocaleType, mergeLocales } from '@univerjs/core'
import { UniverSheetsNodeCorePreset } from '@univerjs/preset-sheets-node-core'
import sheetsNodeCoreEnUS from '@univerjs/preset-sheets-node-core/locales/en-US'
import { TARGET_FUNCTIONS } from './target'
import { ClientFormulaIntegration } from './client'
import { createOssUniver } from './univerTestHarness'

// Fixture columns (see target.ts sample formulas):
//   A: header + 1..4   B: text labels   E: cashflows for IRR   F: has a mode
const FIXTURE = {
  id: 'wb-audit',
  sheets: {
    s1: {
      id: 's1',
      name: 'Data',
      rowCount: 1200,
      columnCount: 40,
      cellData: {
        0: { 0: { v: 'n' }, 1: { v: 'label' }, 4: { v: 'cash' }, 5: { v: 5 } },
        1: { 0: { v: 1 }, 1: { v: 'a' }, 4: { v: -100 }, 5: { v: 5 } },
        2: { 0: { v: 2 }, 1: { v: 'b' }, 4: { v: 60 }, 5: { v: 7 } },
        3: { 0: { v: 3 }, 1: { v: 'c' }, 4: { v: 60 } },
        4: { 0: { v: 4 }, 1: { v: 'd' } },
      },
    },
  },
}

describe('formula coverage audit (real engine)', () => {
  it('evaluates the curated target set without #NAME? regressions', async () => {
    const { univer, univerAPI } = createOssUniver({
      locale: LocaleType.EN_US,
      locales: { [LocaleType.EN_US]: mergeLocales(sheetsNodeCoreEnUS) },
      presets: [UniverSheetsNodeCorePreset()],
    })
    try {
      univerAPI.createWorkbook(FIXTURE)
      const sheet = univerAPI.getActiveWorkbook()!.getActiveSheet()!
      const client = ClientFormulaIntegration.fromUniver(univerAPI)
      const activation = client.activate()

      // Write every target formula into column H, SIX rows apart — dynamic-
      // array results need spill room below or they'd #SPILL! on the next
      // formula (which is a layout collision, not a coverage miss). One
      // calculation pass covers all of them.
      TARGET_FUNCTIONS.forEach((t, i) => {
        activation.setFormula(sheet.getRange(i * 6, 7, 1, 1), t.formula)
      })
      await activation.recalculate(30000)

      const unsupported: string[] = []
      const errored: { name: string; value: unknown }[] = []
      TARGET_FUNCTIONS.forEach((t, i) => {
        const v = sheet.getRange(i * 6, 7, 1, 1).getValue()
        const s = v == null ? '' : String(v)
        if (s === '#NAME?') unsupported.push(t.name)
        // Other error values can be legitimate (our sample args might not
        // match an engine's stricter parsing) — reported, not failed.
        else if (s.startsWith('#') && t.name !== 'NA') errored.push({ name: t.name, value: s })
      })

      // The published matrix, in one honest log line each.

      console.log(
        `[formula-audit] targets=${TARGET_FUNCTIONS.length} supported=${TARGET_FUNCTIONS.length - unsupported.length} unsupported=[${unsupported.join(', ')}] argErrors=[${errored.map((e) => `${e.name}:${e.value}`).join(', ')}]`,
      )

      // Baseline gate: fail on regression below the audited support level.
      // Update BASELINE_UNSUPPORTED deliberately (with a roadmap note) when
      // the target list grows, never to paper over an engine downgrade.
      const BASELINE_UNSUPPORTED: string[] = []
      expect(unsupported.sort()).toEqual(BASELINE_UNSUPPORTED.sort())
      // With spill room and valid samples, argument errors are regressions too.
      expect(errored).toEqual([])
      // Prove real computation happened (not a silent nothing-pass): SUM of
      // the fixture column is a known value.
      const sumIdx = TARGET_FUNCTIONS.findIndex((t) => t.name === 'SUM')
      expect(sheet.getRange(sumIdx * 6, 7, 1, 1).getValue()).toBe(10)
      activation.dispose()
    } finally {
      univer.dispose()
    }
  }, 60000)
})
