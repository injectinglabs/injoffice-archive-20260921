// Package collab is an in-process collaboration hub for InjOffice artifacts.
//
// It implements docs/COLLABORATION-PROTOCOL.md: join/leave/presence, a linear
// operation log with STALE_BASE, catch-up (op.since + reset), and file.changed
// fan-out. Artifact IDs are opaque strings; the hub never resolves them as
// filesystem paths. Authentication and authorization belong to the host, which
// must derive user identity before constructing a Conn. The hub assigns
// client IDs and never reads user_id or client_id from a client payload.
//
// The default store is in-memory and is forgotten when a room empties. A host
// may plug in a durable OpStore later (sqlite/postgres) without changing the
// hub API. HTTP + SSE for a local two-browser demo lives in injoffice-server.
// This package does not serve HTTP or start a process.
package collab
