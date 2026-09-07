# Agent change sets

InjOffice provides a provider-independent execution boundary for document agents. A model, deterministic rule, or human can propose plain operations; InjOffice inspects the target, creates an isolated change set, validates it, applies an approved commit against an expected revision, and verifies the output.

```text
inspect -> plan -> preview + diff -> validate -> approve -> commit -> verify
```

The core packages do not host a model, proxy model traffic, keep API keys, or require an InjOffice service.

## Architecture

`@injoffice/agent-tools` defines the common session, change-set, evidence, limits, refusal, registry, and adapter contracts. `@injoffice/agent-office` provides format adapters. Applications pass an adapter directly or register adapters in `AgentAdapterRegistry`; they also provide the model or other proposal source and the approval policy.

An adapter is responsible for:

1. Reporting exact operations supported for an artifact.
2. Returning bounded inspections tied to a revision and fingerprint.
3. Translating supported semantic operations into native format mutations.
4. Producing an isolated preview and structured diff without changing the source.
5. Refusing unsupported, ambiguous, stale, or policy-blocked work.
6. Applying an approved plan atomically where the underlying engine permits it.
7. Reopening the output and returning format-specific verification evidence.

## Safety contract

- Agents discover capabilities instead of assuming Office parity.
- Reads and operation batches have host-configured limits.
- A change set is immutable and records its source revision.
- Commit requires a matching expected revision and an idempotency key.
- Destructive operations can require an explicit host confirmation callback.
- Structured issues include stable codes and JSON paths.
- A refusal returns no partial output.
- Verification is separate from successful mutation and carries output evidence.

The host still owns identity, authorization, tenant isolation, durable storage, malware scanning, model policy, and the user approval experience.

## Formats

The workflow is format-neutral, but capabilities are not. XLSX, DOCX, PPTX, and PDF adapters expose only the operations that their current native engines can preserve and verify. Unsupported macros, encrypted packages, embedded objects, ambiguous targets, or out-of-scope formatting must be preserved untouched or refused.

New formats implement the same adapter contract. They do not require changes to an agent loop or provider integration.

## Non-goals

This layer is not:

- a built-in chat application or autonomous agent;
- a required hosted backend or model proxy;
- an abstraction that claims every document supports the same mutations;
- a replacement for application authorization or human-review policy;
- permission to silently reconstruct unsupported native files.

## Playground proof

The `#/agent` playground route demonstrates capability discovery, a bounded inspection, a deterministic local operation proposal, isolated preview and diff, validation, explicit human approval, atomic commit, and output verification. No model SDK or network request is used. Its refusal mode proves that unsupported operations stop before write.
