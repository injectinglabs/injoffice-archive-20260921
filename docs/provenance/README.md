# Provenance records

This directory contains machine-readable dependency provenance used by the
repository's license, attribution, and reproducibility checks.

- `npm-license-evidence.json` supplements missing lockfile license metadata only
  where an exact version and npm integrity digest have been reviewed. The
  dependency gate still reports every unresolved package as `NOASSERTION`.
