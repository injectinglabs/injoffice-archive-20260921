# Fixture specs

One declarative JSON spec lives under the `docx`, `pptx`, or `xlsx`
subdirectory for every compatibility fixture. Specs are the source of truth;
run `go run ./cmd/corpusgen` from `go/officecompat` to reproduce packages,
expected-native-JSON sidecars, and `corpus/manifest.json`.

All payloads must be redistributable and carry stable provenance. Prefer exact
hand-authored OOXML and tiny public-domain payloads. `text`, `base64`, and
bounded `repeat` entries are mutually exclusive. The latter exists only for
small-on-disk adversarial compression cases.
