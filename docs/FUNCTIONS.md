# Formula function matrix

Method: not a list comparison — every function below is EVALUATED in the real
Univer engine, headless (`packages/formulas/src/audit.test.ts`, node preset),
and the audit fails CI on any regression: a `#NAME?`, an argument error on the
valid samples, or a wrong SUM proves-computation check.

## Result (2026-08-21, Univer OSS 0.25.1)

- **Declared function inventory: 535** across 15 categories (math 83,
  statistical 113, financial 55, engineering 57, text 51, lookup 37,
  compatibility 38, date 27, information 25, logical 19, database 12, cube 7,
  meta 7, web 3, array 2).
- **Curated high-value target set: 142/142 computing** — zero `#NAME?`, zero
  argument errors. Includes the modern set competitors miss: `LAMBDA`, `LET`,
  `XLOOKUP`/`XMATCH`, `TEXTSPLIT`/`TEXTBEFORE`/`TEXTAFTER`, `IFS`/`SWITCH`,
  dynamic arrays (`UNIQUE`, `SORT`, `FILTER`, `SEQUENCE`) with real spill
  behavior, plus the aggregation, text, date, lookup, and core financial
  families.

InjOffice sheets compute through Univer OSS's engine: 535 declared, with
the 142 that matter proven end-to-end and gated against regression.

`ClientFormulaIntegration` is the host-facing bridge for this audited
vocabulary. The conformance test writes all 142 formulas through its worksheet
range API and explicitly runs the public Univer calculation facade. Custom
host functions are registered in deterministic name order, capability/config
gated, transactionally rolled back on failure, and unregistered in reverse
order on disposal. Audited built-in names cannot be overridden accidentally.

## Gap-closing

No gaps found in the curated set. When one appears (a user/agent formula
hitting `#NAME?`), the path is `univerAPI.getFormula().registerFunction(name,
fn, description)` through `ClientFormulaIntegration` — add the function to
`TARGET_FUNCTIONS` first (the audit turns red), then register the independently
implemented host function (it turns green).

## Leftovers (tracked in ROADMAP Phase 7)

- Agent formula tooling: NL→formula validated against the live workbook,
  explain-this-formula, plain-language error diagnostics — lands with the
  agent-integration phase.
- Golden value-corpus (expected values per function, not just non-error) —
  grow it as formulas become user-visible surface.
