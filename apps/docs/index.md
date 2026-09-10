# Build document workflows with InjOffice

InjOffice is a set of open-source TypeScript packages and Go engines for spreadsheets, documents, presentations, and PDFs. Use the packages independently, add an editor when you need one, or expose supported operations to an AI agent through a reviewable change-set workflow.

These are the developer docs. The [interactive demo](https://injoffice.com/) is a separate application; reading the docs does not load its editors, sample files, or native engines.

## Start with an outcome

| I want to… | Start here |
| --- | --- |
| Make a first file transformation | [Rotate a PDF in a few lines](getting-started/quickstart) |
| Edit a spreadsheet without uploading it | [Browser-local XLSX guide](guides/spreadsheets) |
| Connect a document editor to guarded saves | [Native DOCX guide](guides/documents) |
| Inspect an existing deck or author a new one | [PPTX guide](guides/presentations) |
| Let an agent propose reviewable changes | [Agent quickstart](agents/quickstart) |
| Choose libraries for an existing application | [Package guide](getting-started/packages) |

## A consistent document boundary

For native Office files, retain the original archive, extract a versioned projection, and apply supported mutations against the exact source revision. A rendered page or editor snapshot is not a replacement for the original file.

```text
Original bytes → validated projection → supported mutation → new bytes → readback
```

For agent workflows, insert bounded inspection, a proposed change set, and your application's approval policy before the write. InjOffice does not include an LLM, require API keys, or choose a provider for you.

## What is included

- Task-oriented guides with source-imported TypeScript examples.
- Source-derived reference pages for every TypeScript package and Go module.
- Native file, collaboration, and agent protocol documentation.
- Browser/Node boundaries, error recovery, and production responsibilities.

## What to expect

This documentation tracks the repository's `main` branch. An npm version can lag behind it. Confirm availability and exports before installing; use the [release reference](reference/generated/contracts/public-release) and [source checkout instructions](getting-started/quickstart#from-source) when necessary.

InjOffice implements bounded capabilities, not complete Microsoft Office or Google Workspace parity. Start with the [support matrix](getting-started/support), and always honor runtime refusals and unsupported-content records.
