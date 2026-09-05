# PowerPoint-authored PPTX fixture

This is an unmodified presentation published by Microsoft in the
`microsoft/community-content` repository. Its extended properties name
Microsoft Macintosh PowerPoint as the authoring application. The fixture is
kept byte-for-byte so tests can distinguish actual Office OOXML from synthetic
packages, including the in-repo python-pptx sample.

OfficeDev `office-scripts-docs` is Excel-only and does not publish a
PowerPoint-authored PPTX.

## `attendee-survey-qr.pptx`

- Source: <https://github.com/microsoft/community-content/blob/40bc19c8261b158ff92b0ef26b35d4ffcf570458/archive/SeasonOfAI-S2-Copilots/SeasonOfAI-AttendeeSurveyQR-Slide.pptx>
- Upstream commit: `40bc19c8261b158ff92b0ef26b35d4ffcf570458`
- Upstream Git blob: `79ff3f21150cee5f6edd146e403fd8cbd2495aba`
- Local SHA-256: `b4a503d90634e117ca53fb62d6aaf0657b78bfe0269b7314a313077a4899e095`
- License: [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/)
- Changes: none

This Microsoft Macintosh PowerPoint 16.0000 single-slide community sample
supplies a real Office-authored PresentationML package. Native extract is
fail-closed against its unmodeled picture fill; the fixture is provenance
authority, not a claim that current native PPTX projection covers the slide.
The v1 completion-matrix protocol intentionally does not bind it as
`kind: office-export`; v2 and v3 do.
