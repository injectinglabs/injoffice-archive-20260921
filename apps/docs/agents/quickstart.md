# Your first agent workflow

Start without an LLM. A deterministic proposal is enough to exercise the real document boundary: inspect capabilities, build a revision-bound plan, show a preview/diff, obtain approval, commit, and verify replacement bytes.

## Install the boundaries you need

```sh
npm install @injoffice/agent-tools @injoffice/agent-office @injoffice/pdf
```

Use only the `/pdf` adapter entry for this example. Native Office adapters have additional host requirements; adding a provider SDK does not supply those callbacks.

## Prepare a reviewable PDF change

<<< @/examples/pdf-agent.ts

Call `preparePdfRotation` with actual PDF bytes. It returns the exact plan, preview, diff, and a host-only commit closure. Calling it does **not** save a file or approve the proposal.

In your UI, render the returned review information. Only your trusted review handler should call `commitFromReview` with the approved plan ID and a stable idempotency key. Persist or download the returned bytes after the verification check succeeds.

## The approval boundary

The plan ID comparison prevents accidentally applying approval to a different plan; it is **not authentication**. The application must check the authenticated principal's right to modify this artifact and perform the explicit review action. A model can read plan IDs, so never treat possession of an ID as approval.

The example deliberately does not expose the commit closure as a model tool. Its PDF rotation is classified as non-destructive by the adapter, but the host still gates it behind review. Operations advertised as destructive additionally require the session's `confirmDestructive` hook. See [approval and verification](safety).

## Add a model later, if you want

Replace only the proposal producer. Validate the model's structured output against the discovered operation schemas, enforce limits, and keep the review/commit boundary unchanged. Do not let a proposal supply approval or choose host credentials.

The [public demo](https://injoffice.com/) uses simulated proposals with real bounded document operations across its four formats. This docs example is independent of that application and calls no model service.
