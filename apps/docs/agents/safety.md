# Approval and verification

An agent proposal is untrusted input. InjOffice provides bounded operations and revision checks; the host still decides who may read or modify each artifact.

## Bind approval to the exact plan

The review action should bind the authenticated reviewer, artifact ID, source revision/fingerprint, complete change-set identity, and allowed operation set. Expire that approval when the source or plan changes.

Never accept `confirmation: "approved"` from model output. Do not expose an approval-store mutation as a model tool. A `confirmDestructive` callback reads private host approval state; it does not delegate authorization to the proposal.

Enforce application authorization for **every** write. Adapter flags such as `requiresConfirmation` add a guard for destructive capabilities; they are not a complete access-control policy for non-destructive operations.

## Distinguish the outcomes

| State | Meaning | Host behavior |
| --- | --- | --- |
| Validation/refusal | The plan cannot be applied as requested | Keep the original; show issues |
| Stale revision | Source changed after inspection or review | Re-read, re-plan, and request fresh approval |
| Commit verified | Replacement was written and adapter checks passed | Apply your additional acceptance policy before release |
| Write completed, verification failed | A replacement exists but is not verified | Withhold verified download; retain evidence and source |
| Unknown transport outcome | The response was lost or timed out | Reconcile/retry with the same durable idempotency key |

Do not relabel a post-write verification failure as “nothing changed.” Do not automatically replay an old plan against a newly fetched source revision.

## Idempotency and concurrency

The in-memory session deduplicates its own commits. Persistent or multi-process hosts need a durable idempotency store and atomic revision checks at the actual write boundary. A fresh process cannot rely on another process's session memory.

Keep the idempotency key stable for retries of one logical commit and bind it to the same request. Use a new plan and key for a new logical change.

## Bound model context and tool access

Use capability schemas and bounded reads instead of dumping entire archives into prompts. Enforce document size, operation count, read byte/item budgets, and timeouts. If a real model is introduced, file disclosure, provider policy, credential handling, and prompt-injection defenses remain host responsibilities.

The static demo's no-model behavior is not a claim that an arbitrary downstream model integration is safe.

## Verification limits

Different adapters verify different properties. Fresh parsing proves structural readability, not every intended semantic or visual guarantee. Compare actual requested changes and important preservation evidence using the exact committed bytes. Keep [format-specific limitations](../getting-started/support) visible to users.
