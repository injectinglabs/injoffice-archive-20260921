package main

import "testing"

func TestReferenceIndependenceIsExplicit(t *testing.T) {
	dir, spec := fixture(t)
	legacy, err := execute(t, dir, spec)
	if err != nil || legacy.Cases[0].ReferenceKind != "unclassified" {
		t.Fatalf("legacy source implied independence: %+v %v", legacy, err)
	}
	spec.Version = 2
	for _, kind := range []string{"analytical-oracle", "self-regression", "external-office-export"} {
		spec.Cases[0].ReferenceKind = kind
		spec.Cases[0].ReferenceLicense = "Original test fixture"
		spec.Cases[0].ReferenceProvenance = "Reviewed local test inputs; test metadata, not an Office fidelity result"
		result, err := execute(t, dir, spec)
		if err != nil || !result.Passed || result.Cases[0].ReferenceKind != kind {
			t.Fatalf("%s: %+v %v", kind, result, err)
		}
	}
}

func TestV2RefusesUnclassifiedOrMisleadingReferences(t *testing.T) {
	for _, mutation := range []func(*visualCase){
		func(c *visualCase) { c.ReferenceKind = "" },
		func(c *visualCase) { c.ReferenceKind = "perfect-office-parity" },
		func(c *visualCase) { c.ReferenceLicense = "" },
		func(c *visualCase) { c.ReferenceProvenance = "" },
		func(c *visualCase) { c.CandidateRenderer = c.ReferenceRenderer },
	} {
		dir, spec := fixture(t)
		spec.Version = 2
		spec.Cases[0].ReferenceKind = "external-office-export"
		spec.Cases[0].ReferenceLicense = "Original fixture"
		spec.Cases[0].ReferenceProvenance = "Exported manually with recorded application version and settings"
		mutation(&spec.Cases[0])
		if _, err := execute(t, dir, spec); err == nil {
			t.Fatal("accepted misleading or incomplete reference metadata")
		}
	}
}
