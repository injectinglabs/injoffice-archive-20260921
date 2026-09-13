# XLSX XML UTF-8 signatures

Native XLSX extraction accepts one UTF-8 byte-order mark (`EF BB BF`) at byte zero of supported XML parts. XML 1.0 §4.3.3 identifies this optional UTF-8 prefix as an encoding signature, not markup or character data: https://www.w3.org/TR/xml/#charencoding .

Preflight and the non-offset workbook-dialect/relationship readers skip exactly one prefix in their decoder view. The shared root reader validates the initial character-data token using its original byte offset; it does not shift the input reader. Original ZIP and part bytes, source hashes, and offsets used for mutation remain unchanged. U+FEFF within element text remains text.

Repeated or noninitial signatures outside the root, misplaced declarations, malformed encoding, unsupported encodings, DTDs, invalid namespaces, and existing package/XML limits retain their checks. This is not general UTF-16 support or a relaxation of style authority.

Public Extract/Inspect qualification against unchanged ClosedXML fixtures demonstrates that `Examples/Styles/UsingRichText.xlsx` now opens. `Other/InlinedRichText/ChangeRichText/inputfile.xlsx` progresses past its signature but still refuses because its `cellStyleXfs` lacks the count required by the existing style profile. These are library fixtures, not independent Excel visual references; source files and browser evidence stay in the external local corpus.
